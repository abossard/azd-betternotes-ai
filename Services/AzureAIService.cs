using Azure;
using Azure.AI.FormRecognizer.DocumentAnalysis;
using Azure.Core;
using Azure.Identity;
using DocumentFormat.OpenXml.Packaging;
using PdfSharpCore.Drawing;
using PdfSharpCore.Pdf;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BetterNotes.Services
{
    public class AzureAIService
    {
        private readonly string _endpoint;
        private readonly string _key;
        private readonly TokenCredential _credential = null!;
        private readonly bool _useManagedIdentity;
        private readonly bool _isConfigured;
        private readonly ILogger<AzureAIService> _logger;
        private readonly DocumentAnalysisClient _documentClient = null!;

        public AzureAIService(IConfiguration configuration, ILogger<AzureAIService> logger)
        {
            _endpoint = configuration["AzureAI:Endpoint"];
            _key = configuration["AzureAI:Key"];
            _logger = logger;

            _isConfigured = !string.IsNullOrWhiteSpace(_endpoint);
            if (!_isConfigured)
            {
                _logger.LogWarning("AzureAI:Endpoint is not configured. Azure AI analysis will be unavailable.");
                return;
            }

            var endpointUri = new Uri(_endpoint);

            // Use Managed Identity if no key is provided
            _useManagedIdentity = string.IsNullOrEmpty(_key);
            if (_useManagedIdentity)
            {
                _credential = new DefaultAzureCredential();
                _documentClient = new DocumentAnalysisClient(endpointUri, _credential);
                _logger.LogInformation("Using Managed Identity for authentication");
            }
            else
            {
                _documentClient = new DocumentAnalysisClient(endpointUri, new AzureKeyCredential(_key));
                _logger.LogInformation("Using API Key for authentication");
            }
        }

        public async Task<string> AnalyzeFileAsync(Stream fileStream, string fileName)
        {
            _logger.LogInformation("Starting analysis for file: {FileName}", fileName);
            _logger.LogInformation("Authentication mode: {Mode}", _useManagedIdentity ? "Managed Identity" : "API Key");
            _logger.LogInformation("Endpoint: {Endpoint}", _endpoint);

            if (!_isConfigured)
            {
                return "Error: Azure AI endpoint is not configured. Set AzureAI:Endpoint to enable analysis.";
            }

            var extension = Path.GetExtension(fileName).ToLowerInvariant();
            var workingStream = new MemoryStream();

            try
            {
                var isSupported = extension switch
                {
                    ".pdf" => true,
                    ".docx" => true,
                    ".jpg" or ".jpeg" => true,
                    ".png" => true,
                    ".bmp" => true,
                    ".tiff" or ".tif" => true,
                    _ => false
                };

                if (!isSupported)
                {
                    return $"Error: Unsupported file type '{extension}'. Supported types for analysis are PDF, DOCX, JPG, PNG, BMP, and TIFF.";
                }

                if (fileStream.CanSeek)
                {
                    fileStream.Position = 0;
                }
                await fileStream.CopyToAsync(workingStream);
                workingStream.Position = 0;

                if (extension == ".docx")
                {
                    var handwrittenPdfStream = TryCreatePdfFromLikelyHandwrittenDocx(workingStream, fileName);
                    if (handwrittenPdfStream != null)
                    {
                        using (handwrittenPdfStream)
                        {
                            try
                            {
                                var handwrittenResult = await AnalyzeWithModelAsync(handwrittenPdfStream, "prebuilt-read");
                                if (!string.IsNullOrWhiteSpace(handwrittenResult.Content))
                                {
                                    _logger.LogInformation("Used auto DOCX-to-PDF OCR path for file: {FileName}", fileName);
                                    return "=== Handwritten Notes (Auto PDF OCR) ===\n" + ExtractTextFromAnalyzeResult(handwrittenResult);
                                }
                            }
                            catch (RequestFailedException handwrittenEx)
                            {
                                _logger.LogWarning(handwrittenEx, "Auto DOCX-to-PDF OCR path failed for file: {FileName}; continuing with standard DOCX analysis.", fileName);
                            }
                        }
                    }
                }

                var result = await AnalyzeWithModelAsync(workingStream, "prebuilt-document");

                if (string.IsNullOrWhiteSpace(result.Content))
                {
                    _logger.LogWarning("Azure AI returned an empty analysis result for file: {FileName}", fileName);
                    return "Error: Empty response from Azure AI.";
                }

                return ExtractTextFromAnalyzeResult(result);
            }
            catch (RequestFailedException ex) when (ex.Status == 400)
            {
                _logger.LogError(ex, "Document Intelligence returned BadRequest for file: {FileName}. ErrorCode: {ErrorCode}", fileName, ex.ErrorCode);

                if (extension == ".docx")
                {
                    var retryPdfStream = TryCreatePdfFromDocxImages(workingStream, fileName, useHeuristic: false);
                    if (retryPdfStream != null)
                    {
                        using (retryPdfStream)
                        {
                            try
                            {
                                var retryResult = await AnalyzeWithModelAsync(retryPdfStream, "prebuilt-read");
                                if (!string.IsNullOrWhiteSpace(retryResult.Content))
                                {
                                    _logger.LogInformation("Recovered DOCX BadRequest by converting embedded images to PDF for file: {FileName}", fileName);
                                    return "=== Handwritten Notes (Auto PDF OCR) ===\n" + ExtractTextFromAnalyzeResult(retryResult);
                                }
                            }
                            catch (RequestFailedException retryEx)
                            {
                                _logger.LogWarning(retryEx, "DOCX BadRequest recovery via PDF OCR failed for file: {FileName}", fileName);
                            }
                        }
                    }

                    return await ExtractWordTextFallbackAsync(fileStream, fileName, ex.Message);
                }

                return $"Error: Unable to analyze file. Status: BadRequest. Details: {ex.Message}";
            }
            catch (RequestFailedException ex) when (ex.Status == 403)
            {
                _logger.LogError(ex, "Document Intelligence access forbidden for file: {FileName}. ErrorCode: {ErrorCode}", fileName, ex.ErrorCode);
                return "Error: Access forbidden (403). The managed identity may not have the required permissions yet. Please wait 5-10 minutes for role assignments to propagate, then try again.";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error analyzing file: {FileName}", fileName);
                throw;
            }
        }

        private string ExtractTextFromAnalyzeResult(AnalyzeResult analyzeResult)
        {
            var sb = new StringBuilder();

            try
            {
                if (!string.IsNullOrWhiteSpace(analyzeResult.Content))
                {
                    sb.AppendLine("=== Document Content ===");
                    sb.AppendLine(analyzeResult.Content);
                    sb.AppendLine();
                }

                if (analyzeResult.KeyValuePairs.Count > 0)
                {
                    sb.AppendLine("=== Key-Value Pairs ===");
                    foreach (var kvp in analyzeResult.KeyValuePairs)
                    {
                        var keyText = kvp.Key?.Content ?? "";
                        var valueText = kvp.Value?.Content ?? "";
                        sb.AppendLine($"{keyText}: {valueText}");
                    }
                    sb.AppendLine();
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error extracting text from analyze result.");
                sb.AppendLine("Error: Failed to extract all information from the analysis result.");
            }

            return sb.ToString();
        }

        private async Task<AnalyzeResult> AnalyzeWithModelAsync(Stream inputStream, string modelId)
        {
            if (inputStream.CanSeek)
            {
                inputStream.Position = 0;
            }

            var operation = await _documentClient.AnalyzeDocumentAsync(WaitUntil.Completed, modelId, inputStream);
            return operation.Value;
        }

        private MemoryStream TryCreatePdfFromLikelyHandwrittenDocx(MemoryStream docxStream, string fileName)
        {
            return TryCreatePdfFromDocxImages(docxStream, fileName, useHeuristic: true);
        }

        private MemoryStream TryCreatePdfFromDocxImages(MemoryStream docxStream, string fileName, bool useHeuristic)
        {
            try
            {
                if (docxStream.CanSeek)
                {
                    docxStream.Position = 0;
                }

                using var clone = new MemoryStream(docxStream.ToArray());
                using var wordDoc = WordprocessingDocument.Open(clone, false);

                var plainText = wordDoc.MainDocumentPart?.Document?.Body?.InnerText ?? string.Empty;
                var imageParts = wordDoc.MainDocumentPart?.ImageParts?.ToList() ?? new List<ImagePart>();
                var imageBytes = imageParts
                    .Where(part => IsSupportedImagePart(part.ContentType))
                    .Select(ReadImagePartBytes)
                    .Where(bytes => bytes.Length > 0)
                    .ToList();

                if (imageBytes.Count == 0)
                {
                    return null;
                }

                var likelyHandwrittenDocx = plainText.Length < 800;
                if (useHeuristic && !likelyHandwrittenDocx)
                {
                    return null;
                }

                var pdfBytes = BuildPdfFromImages(imageBytes);
                _logger.LogInformation("Detected likely handwritten DOCX for {FileName} (text length: {TextLength}, images: {ImageCount}). Generated PDF for OCR.", fileName, plainText.Length, imageBytes.Count);
                return new MemoryStream(pdfBytes);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Failed to build PDF from DOCX embedded images for file: {FileName}", fileName);
                return null;
            }
        }

        private static bool IsSupportedImagePart(string contentType)
        {
            return contentType == "image/jpeg"
                || contentType == "image/png"
                || contentType == "image/bmp"
                || contentType == "image/tiff";
        }

        private static byte[] ReadImagePartBytes(ImagePart imagePart)
        {
            using var imageStream = imagePart.GetStream();
            using var buffer = new MemoryStream();
            imageStream.CopyTo(buffer);
            return buffer.ToArray();
        }

        private static byte[] BuildPdfFromImages(IReadOnlyList<byte[]> imageBytes)
        {
            using var output = new MemoryStream();
            using var document = new PdfDocument();

            foreach (var bytes in imageBytes)
            {
                using var image = XImage.FromStream(() => new MemoryStream(bytes));
                var page = document.AddPage();
                page.Width = image.PointWidth;
                page.Height = image.PointHeight;

                using var graphics = XGraphics.FromPdfPage(page);
                graphics.DrawImage(image, 0, 0, page.Width, page.Height);
            }

            document.Save(output, false);
            return output.ToArray();
        }

        private async Task<string> ExtractWordTextFallbackAsync(Stream originalStream, string fileName, string azureErrorDetails)
        {
            try
            {
                if (originalStream.CanSeek)
                {
                    originalStream.Position = 0;
                }

                using var copy = new MemoryStream();
                await originalStream.CopyToAsync(copy);
                copy.Position = 0;

                using var wordDoc = WordprocessingDocument.Open(copy, false);
                var text = wordDoc.MainDocumentPart?.Document?.Body?.InnerText;

                if (string.IsNullOrWhiteSpace(text))
                {
                    return "Error: Azure Document Intelligence could not process this Word file, and no extractable text was found. This often means handwriting is stored as Word ink/drawing objects rather than images. Please export to PDF and upload the PDF for OCR.";
                }

                _logger.LogWarning("Azure analysis failed for DOCX file {FileName}; returned local OpenXML text extraction fallback.", fileName);
                return $"=== Word Text (Fallback Extraction) ===\n{text}\n\n[Note] Azure analysis failed with: {azureErrorDetails}";
            }
            catch (Exception fallbackEx)
            {
                _logger.LogError(fallbackEx, "DOCX fallback extraction also failed for file: {FileName}", fileName);
                return $"Error: Unable to analyze file. Status: BadRequest. Details: {azureErrorDetails}";
            }
        }
    }
}
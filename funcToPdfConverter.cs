using System.Net;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;
using System.Threading.Tasks;
using System;
using Azure.Storage.Blobs;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;
using Syncfusion.OfficeChart;
using System.IO;

namespace AnGiunta.DocToPdf
{
    public static class funcToPdfConverter
    {
        [Function("HttpTriggerConverter")]
        public static async Task<HttpResponseData> HttpTriggerConverter(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post")]
            HttpRequestData req,
            FunctionContext executionContext)
        {
            var logger = executionContext.GetLogger("PdfToDocConverter");
            
            // Open the file and upload its data
            var convertedItem = SyncfusionConvertToPDF(req.Body);
            var blobClient = await GetBlobClient(logger);
            await blobClient.UploadAsync(convertedItem);

            var response = req.CreateResponse(HttpStatusCode.OK);
            response.Headers.Add("Content-Type", "text/plain; charset=utf-8");
            response.WriteString("Welcome to Azure Functions!");
            return response;
        }

        [Function("BlobTriggerConverter")]
        public static async void BlobTriggerConverter(
            [BlobTrigger("toconvert/{name}")] byte[] input, 
            string name,
            FunctionContext executionContext)
        {
            var logger = executionContext.GetLogger("PdfToDocConverter");
            Stream stream = new MemoryStream(input);

            // Open the file and upload its data
            var convertedItem = SyncfusionConvertToPDF(stream);
            var blobClient = await GetBlobClient(logger,name);
            await blobClient.UploadAsync(convertedItem);
        }

        private static async Task<BlobClient> GetBlobClient(ILogger logger, string fileName="")
        {
            string connectionString = Environment.GetEnvironmentVariable("AzureWebJobsStorage");
            
            // Create a BlobServiceClient object which will be used to create a container client
            BlobServiceClient blobServiceClient = new BlobServiceClient(connectionString);
            logger.LogInformation($"ConnectionString:{connectionString}");

            string containerName = "converted";
            logger.LogInformation($"containername {containerName}");

            // Create the container and return a container client object
            BlobContainerClient containerClient = blobServiceClient.GetBlobContainerClient(containerName);
            await containerClient.CreateIfNotExistsAsync();
            
            if(string.IsNullOrEmpty(fileName)) fileName = "converted" + Guid.NewGuid().ToString() + ".pdf";
            else 
            {
                var splitted = fileName.Split('.');
                fileName="";
                for(int i=0; i<splitted.Length-1;i++)
                    fileName+=splitted[i];
                fileName+=".pdf";
            }
                logger.LogInformation($"filename:{fileName}");

            // Get a reference to a blob
            return containerClient.GetBlobClient(fileName);
            
        }
        private static Stream SyncfusionConvertToPDF(Stream input)
        {
            WordDocument wordDocument = new WordDocument(input, Syncfusion.DocIO.FormatType.Docx);
            
            //Instantiation of DocIORenderer for Word to PDF conversion
            DocIORenderer render = new DocIORenderer();
            
            //Sets Chart rendering Options.
            render.Settings.ChartRenderingOptions.ImageFormat =  ExportImageFormat.Jpeg;
            
            //Converts Word document into PDF document
            PdfDocument pdfDocument = render.ConvertToPDF(wordDocument);
            //Releases all resources used by the Word document and DocIO Renderer objects
            render.Dispose();
            wordDocument.Dispose();
            //Saves the PDF file
            MemoryStream outputStream = new MemoryStream();
            pdfDocument.Save(outputStream);
            //Closes the instance of PDF document object
            pdfDocument.Close();
            outputStream.Seek(0,SeekOrigin.Begin);
            return outputStream;
        }
    }
    
}

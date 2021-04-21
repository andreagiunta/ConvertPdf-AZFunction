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
    public static class PdfToDocConverter
    {
        [Function("PdfToDocConverter")]
        public static async Task<HttpResponseData> Run([HttpTrigger(AuthorizationLevel.Anonymous, "get", "post")]
            HttpRequestData req,
            FunctionContext executionContext)
        {
            var logger = executionContext.GetLogger("PdfToDocConverter");
            logger.LogInformation("C# HTTP trigger function processed a request.");
            
            string connectionString = Environment.GetEnvironmentVariable("AzureWebJobsStorage");
            // Create a BlobServiceClient object which will be used to create a container client
            BlobServiceClient blobServiceClient = new BlobServiceClient(connectionString);
            logger.LogInformation($"ConnectionString:{connectionString}");

            //Create a unique name for the container
            string containerName = "converted";
            logger.LogInformation($"containername {containerName}");
            // Create the container and return a container client object
            BlobContainerClient containerClient = blobServiceClient.GetBlobContainerClient(containerName);
            await containerClient.CreateIfNotExistsAsync();
            string fileName = "test" + Guid.NewGuid().ToString() + ".pdf";
            logger.LogInformation($"filename:{fileName}");

            // Get a reference to a blob
            BlobClient blobClient = containerClient.GetBlobClient(fileName);

            // Open the file and upload its data
            //await blobClient.UploadAsync(req.Body);
            logger.LogInformation($"BodyLenght:{req.Body.Length}");

            WordDocument wordDocument = new WordDocument(req.Body, Syncfusion.DocIO.FormatType.Docx);
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
            await blobClient.UploadAsync(outputStream);
            



            var response = req.CreateResponse(HttpStatusCode.OK);
            response.Headers.Add("Content-Type", "text/plain; charset=utf-8");
            response.WriteString("Welcome to Azure Functions!");
            return response;
        }
    }
}

using Google.Cloud.Functions.Framework;
using Google.Cloud.Storage.V1;
using GrapeCity.Documents.Excel;
using Microsoft.AspNetCore.Http;
using System.IO;
using System.Threading.Tasks;

namespace ExcelExportCloudHttpFunction2
{
    public class Function : IHttpFunction
    {
        /// <summary>
        /// Logic for your function goes here.
        /// </summary>
        /// <param name="context">The HTTP context, containing the request and the response.</param>
        /// <returns>A task representing the asynchronous operation.</returns>
        public async Task HandleAsync(HttpContext context)
        {
            HttpRequest request = context.Request;
            string name = request.Query["name"].ToString();

            string Message = string.IsNullOrEmpty(name)
                ? "Hello, World!!"
                : $"Hello, {name}!!";

            // トライアル版または製品版のライセンスキーを設定
            //Workbook.SetLicenseKey("");

            Workbook workbook = new Workbook();
            workbook.Worksheets[0].Range["A1"].Value = Message;

            using MemoryStream outputstream = new MemoryStream();
            StorageClient sc = StorageClient.Create();
            workbook.Save(outputstream, SaveFileFormat.Xlsx);
            await sc.UploadObjectAsync("diodocs-export", "Result.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", outputstream);
        }
    }
}

using ImportExportFiles.Helper;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.IO;

namespace ImportExportFiles.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class FileController : ControllerBase
    {
        private readonly ILogger<FileController> _logger;

        public FileController(ILogger<FileController> logger)
        {
            _logger = logger;
        }

        [HttpGet("ExportToPreDefinedExcelTemplate")]
        public IActionResult ExportToPreDefinedExcelTemplate()
        {
            TemplateExcel.FillReport("invoice.xlsx", "template.xlsx", DataSample.GetDataSet(), new string[] { "{", "}" });
            //Process.Start("invoice.xlsx");
            var excelName = $"UserList-{DateTime.Now:yyyyMMddHHmmssfff}.xlsx";

            //reopen file for fix bug
            var fs = new FileStream("invoice.xlsx", FileMode.Open, FileAccess.Read);
            return File(fs, "application/octet-stream", excelName);
        }
    }
}
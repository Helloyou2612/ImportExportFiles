using ImportExportFiles.Helper;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.IO;

namespace ImportExportFiles.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class FileController : ControllerBase
    {
        private readonly ILogger<FileController> _logger;
        private readonly IHostingEnvironment _env;

        public FileController(ILogger<FileController> logger, IHostingEnvironment env)
        {
            _logger = logger;
            _env = env;
        }

        [HttpGet("ExportToPreDefinedExcelTemplate")]
        public IActionResult ExportToPreDefinedExcelTemplate()
        {
            TemplateExcel.FillReport("invoice.xlsx", "template.xlsx", DataSample.GetDataSet(), new string[] { "{", "}" });
            //Process.Start("invoice.xlsx");
            var fileName = $"UserList-{DateTime.Now:yyyyMMddHHmmssfff}.xlsx";

            //reopen file for fix bug
            var fs = new FileStream("invoice.xlsx", FileMode.Open, FileAccess.Read);
            return File(fs, "application/octet-stream", fileName);
        }

        [HttpGet("ExportToPivotTableExcel")]
        public IActionResult ExportToPivotTableExcel()
        {
            var webRoot = _env.ContentRootPath;

            var fileName = $"UserList-{DateTime.Now:yyyyMMddHHmmssfff}.xlsx";
            var folderExportPath = Path.Combine(webRoot + "/Files/Export/");
            //the path of the file
            string filePath = Path.Combine(folderExportPath + fileName);
            
            if (!Directory.Exists(folderExportPath))
            {
                Directory.CreateDirectory(folderExportPath);
            }

            //create a new ExcelPackage
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                //create 2 WorkSheets. One for the source data and one for the Pivot table
                ExcelWorksheet worksheetPivot = excelPackage.Workbook.Worksheets.Add("Pivot");
                ExcelWorksheet worksheetData = excelPackage.Workbook.Worksheets.Add("Data");

                //add some source data
                worksheetData.Cells["A1"].Value = "Column A";
                worksheetData.Cells["A2"].Value = "Group A";
                worksheetData.Cells["A3"].Value = "Group B";
                worksheetData.Cells["A4"].Value = "Group C";
                worksheetData.Cells["A5"].Value = "Group A";
                worksheetData.Cells["A6"].Value = "Group B";
                worksheetData.Cells["A7"].Value = "Group C";
                worksheetData.Cells["A8"].Value = "Group A";
                worksheetData.Cells["A9"].Value = "Group B";
                worksheetData.Cells["A10"].Value = "Group C";
                worksheetData.Cells["A11"].Value = "Group D";

                worksheetData.Cells["B1"].Value = "Column B";
                worksheetData.Cells["B2"].Value = "emc";
                worksheetData.Cells["B3"].Value = "fma";
                worksheetData.Cells["B4"].Value = "h2o";
                worksheetData.Cells["B5"].Value = "emc";
                worksheetData.Cells["B6"].Value = "fma";
                worksheetData.Cells["B7"].Value = "h2o";
                worksheetData.Cells["B8"].Value = "emc";
                worksheetData.Cells["B9"].Value = "fma";
                worksheetData.Cells["B10"].Value = "h2o";
                worksheetData.Cells["B11"].Value = "emc";

                worksheetData.Cells["C1"].Value = "Column C";
                worksheetData.Cells["C2"].Value = 299;
                worksheetData.Cells["C3"].Value = 792;
                worksheetData.Cells["C4"].Value = 458;
                worksheetData.Cells["C5"].Value = 299;
                worksheetData.Cells["C6"].Value = 792;
                worksheetData.Cells["C7"].Value = 458;
                worksheetData.Cells["C8"].Value = 299;
                worksheetData.Cells["C9"].Value = 792;
                worksheetData.Cells["C10"].Value = 458;
                worksheetData.Cells["C11"].Value = 299;

                worksheetData.Cells["D1"].Value = "Column D";
                worksheetData.Cells["D2"].Value = 40075;
                worksheetData.Cells["D3"].Value = 31415;
                worksheetData.Cells["D4"].Value = 384400;
                worksheetData.Cells["D5"].Value = 40075;
                worksheetData.Cells["D6"].Value = 31415;
                worksheetData.Cells["D7"].Value = 384400;
                worksheetData.Cells["D8"].Value = 40075;
                worksheetData.Cells["D9"].Value = 31415;
                worksheetData.Cells["D10"].Value = 384400;
                worksheetData.Cells["D11"].Value = 40075;

                //define the data range on the source sheet
                var dataRange = worksheetData.Cells[worksheetData.Dimension.Address];

                //create the pivot table
                var pivotTable = worksheetPivot.PivotTables.Add(worksheetPivot.Cells["B2"], dataRange, "PivotTable");

                //label field
                pivotTable.RowFields.Add(pivotTable.Fields["Column A"]);
                pivotTable.DataOnRows = false;

                //data fields
                var field = pivotTable.DataFields.Add(pivotTable.Fields["Column B"]);
                field.Name = "Count of Column B";
                field.Function = DataFieldFunctions.Count;

                field = pivotTable.DataFields.Add(pivotTable.Fields["Column C"]);
                field.Name = "Sum of Column C";
                field.Function = DataFieldFunctions.Sum;
                field.Format = "0.00";

                field = pivotTable.DataFields.Add(pivotTable.Fields["Column D"]);
                field.Name = "Sum of Column D";
                field.Function = DataFieldFunctions.Sum;
                field.Format = "€#,##0.00";

                //Write the file to the disk
                FileInfo fi = new FileInfo(filePath);
                excelPackage.SaveAs(fi);
            }

            //reopen file for fix bug file is openning
            var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            return File(fs, "application/octet-stream", fileName);
        }
    }
}
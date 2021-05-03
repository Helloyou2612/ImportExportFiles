using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Data;
using System.Drawing;
using System.IO;

namespace ImportExportFiles.Helper
{
    public class EPPlus
    {
        private readonly FileInfo _newFile;
        private readonly FileInfo _templateFile;
        private readonly DataSet _ds;
        private ExcelPackage _xlPackage; 
        public string ErrorMessage;

        public EPPlus(string filePath, string templateFilePath)
        {
            _newFile = new FileInfo(@filePath);
            _templateFile = new FileInfo(@templateFilePath);
            _ds = null;//GetDataTables(); /* DataTables */
            ErrorMessage = string.Empty;
            CreateFileWithTemplate();
        }

        private bool CreateFileWithTemplate()
        {
            try
            {
                ErrorMessage = string.Empty;

                using (_xlPackage = new ExcelPackage(_newFile, _templateFile))
                {
                    var i = 1;
                    foreach (DataTable dt in _ds.Tables)
                    {
                        AddSheetWithTemplate(_xlPackage, dt, i);
                        i++;
                    }

                    ///* Set title, Author.. */
                    //xlPackage.Workbook.Properties.Title = "Title: Office Open XML Sample";
                    //xlPackage.Workbook.Properties.Author = "Author: Muhammad Mubashir.";
                    ////xlPackage.Workbook.Properties.SetCustomPropertyValue("EmployeeID", "1147");
                    //xlPackage.Workbook.Properties.Comments = "Sample Record Details";
                    //xlPackage.Workbook.Properties.Company = "TRG Tech.";

                    ///* Save */
                    _xlPackage.Save();
                }
                return true;
            }
            catch (Exception ex)
            {
                ErrorMessage = ex.Message.ToString();
                return false;
            }
        }

        /// <summary>
        /// This AddSheet method generates a .xlsx Sheet with your provided Template file, //DataTable and SheetIndex.
        /// </summary>
        public static void AddSheetWithTemplate(ExcelPackage xlApp, DataTable dt, int sheetIndex)
        {
            var sheetName = $"Sheet{sheetIndex.ToString()}";
            ExcelWorksheet worksheet;
            /* WorkSheet */
            if (sheetIndex == 0)
            {
                worksheet = xlApp.Workbook.Worksheets[sheetIndex + 1]; // add a new worksheet to the empty workbook
            }
            else
            {
                worksheet = xlApp.Workbook.Worksheets[sheetIndex]; // add a new worksheet to the empty workbook
            }

            if (worksheet == null)
            {
                worksheet = xlApp.Workbook.Worksheets.Add(sheetName); // add a new worksheet to the empty workbook
            }

            /* Load the datatable into the sheet, starting from cell A1. Print the column names on row 1 */
            worksheet.Cells["A1"].LoadFromDataTable(dt, true);
        }

        private static void AddSheet(ExcelPackage xlApp, DataTable dt, int Index, string sheetName)
        {
            var _sheetName = string.Empty;

            if (string.IsNullOrEmpty(sheetName) == true)
            {
                _sheetName = $"Sheet{Index.ToString()}";
            }
            else
            {
                _sheetName = sheetName;
            }

            /* WorkSheet */
            var worksheet = xlApp.Workbook.Worksheets[_sheetName]; // add a new worksheet to the empty workbook
            if (worksheet == null)
            {
                worksheet = xlApp.Workbook.Worksheets.Add(_sheetName); // add a new worksheet to the empty workbook
            }
            else
            {
            }

            /* Load the datatable into the sheet, starting from cell A1. Print the column names on row 1 */
            worksheet.Cells["A1"].LoadFromDataTable(dt, true);

            var rowCount = dt.Rows.Count;
            var colCount = dt.Columns.Count;

            #region Set Column Type to Date using LINQ.

            /*
                IEnumerable<int> dateColumns = from DataColumn d in dt.Columns
                                               where d.DataType == typeof(DateTime) || d.ColumnName.Contains("Date")
                                               select d.Ordinal + 1;

                foreach (int dc in dateColumns)
                {
                    xlSheet.Cells[2, dc, rowCount + 1, dc].Style.Numberformat.Format = "dd/MM/yyyy";
                }
                */

            #endregion Set Column Type to Date using LINQ.

            #region Set Column Type to Date using LOOP.

            /* Set Column Type to Date. */
            for (var i = 0; i < dt.Columns.Count; i++)
            {
                if ((dt.Columns[i].DataType).FullName == "System.DateTime" && (dt.Columns[i].DataType).Name == "DateTime")
                {
                    //worksheet.Cells[2,4] .Style.Numberformat.Format = "yyyy-mm-dd h:mm"; //OR "yyyy-mm-dd h:mm" if you want to include the time!
                    worksheet.Column(i + 1).Style.Numberformat.Format = "dd/MM/yyyy HH:mm"; //OR "yyyy-mm-dd h:mm" if you want to include the time!
                    worksheet.Column(i + 1).Width = 25;
                }
            }

            #endregion Set Column Type to Date using LOOP.

            //(from DataColumn d in dt.Columns select d.Ordinal + 1).ToList().ForEach(dc =>
            //{
            //    //background color
            //    worksheet.Cells[1, 1, 1, dc].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //    worksheet.Cells[1, 1, 1, dc].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightYellow);

            //    //border
            //    worksheet.Cells[1, dc, rowCount + 1, dc].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            //    worksheet.Cells[1, dc, rowCount + 1, dc].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            //    worksheet.Cells[1, dc, rowCount + 1, dc].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            //    worksheet.Cells[1, dc, rowCount + 1, dc].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            //    worksheet.Cells[1, dc, rowCount + 1, dc].Style.Border.Top.Color.SetColor(System.Drawing.Color.LightGray);
            //    worksheet.Cells[1, dc, rowCount + 1, dc].Style.Border.Right.Color.SetColor(System.Drawing.Color.LightGray);
            //    worksheet.Cells[1, dc, rowCount + 1, dc].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.LightGray);
            //    worksheet.Cells[1, dc, rowCount + 1, dc].Style.Border.Left.Color.SetColor(System.Drawing.Color.LightGray);
            //});

            /* Format the header: Prepare the range for the column headers */
            var cellRange = "A1:" + Convert.ToChar('A' + colCount - 1) + 1;
            using (var rng = worksheet.Cells[cellRange])
            {
                rng.Style.Font.Bold = true;
                rng.Style.Fill.PatternType = ExcelFillStyle.Solid;                      //Set Pattern for the background to Solid
                rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189));  //Set color to dark blue
                rng.Style.Font.Color.SetColor(Color.White);
            }

            /* Header Footer */
            worksheet.HeaderFooter.OddHeader.CenteredText = "Header: Tinned Goods Sales";
            worksheet.HeaderFooter.OddFooter.RightAlignedText = $"Footer: Page {ExcelHeaderFooter.PageNumber} of {ExcelHeaderFooter.NumberOfPages}"; // add the page number to the footer plus the total number of pages
        }
    }// class End.
}
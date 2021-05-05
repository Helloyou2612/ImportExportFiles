using ImportExportFiles.Helper;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.Data;
using System.IO;
using System.Threading.Tasks;

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

        [HttpGet]
        public async Task<IActionResult> Get()
        {
            TemplateExcel.FillReport("invoice.xlsx", "template.xlsx", GetDataSet(), new string[] { "{", "}" });
            //Process.Start("invoice.xlsx");
            var excelName = $"UserList-{DateTime.Now:yyyyMMddHHmmssfff}.xlsx";

            //reopen file for fix bug 
            var fs = new FileStream("invoice.xlsx", FileMode.Open, FileAccess.Read);
            return File(fs, "application/octet-stream", excelName);
        }

        public DataSet GetDataSet()
        {
            var product = new DataTable();
            product.Columns.Add("tenhang", typeof(string));
            product.Columns.Add("soluong", typeof(int));
            product.Columns.Add("dongia", typeof(int));
            product.Rows.Add("Kỷ thuật lập trình C#", 5, 55000);
            product.Rows.Add("Cơ sở dữ liệu và thuật toán", 3, 15000);
            product.Rows.Add("Giáo trình Photoshop", 20, 65000);
            product.Rows.Add("Triết học", 7, 15000);
            product.Rows.Add("Lập trình mạng Cisco", 2, 21000);
            product.Rows.Add("Làm chủ Microsoft Office 2019", 3, 89000);
            product.Rows.Add("Lập trình hướng đối tượng JAVA", 1, 150000);
            product.Rows.Add("Giáo trình Android/IOS", 8, 90000);
            product.TableName = "productdetails";

            var info = new DataTable();
            info.Columns.Add("tencuahang");
            info.Columns.Add("diachi");
            info.Columns.Add("tenkhachhang");
            info.Columns.Add("diachikhachhang");
            info.Columns.Add("ngaythang");
            info.Columns.Add("dienthoai");
            info.Rows.Add("NHÀ SÁCH TIN HỌC VB.NET", "Địa chỉ: 05/27 Trung Thành, Quảng Thành, Châu Đức, BRVT", "Tên khách hàng: Nguyễn Thảo", "Địa chỉ: Biên Hòa - Đồng Nai", "Biên Hòa, Ngày 02 tháng 12 năm 2020", "Điện thoại: 0933.913.122");
            info.TableName = "info";
            var ds = new DataSet();
            ds.Tables.Add(product);
            ds.Tables.Add(info);
            return ds;
        }
    }
}
using System.Data;

namespace ImportExportFiles.Helper
{
    public class DataSample
    {
        public static DataSet GetDataSet()
        {
            var product = new DataTable();
            product.Columns.Add("stt", typeof(int));
            product.Columns.Add("tenhang", typeof(string));
            product.Columns.Add("soluong", typeof(int));
            product.Columns.Add("dongia", typeof(int));
            product.Rows.Add(1, "Kỷ thuật lập trình C#", 5, 55000);
            product.Rows.Add(2, "Cơ sở dữ liệu và thuật toán", 3, 15000);
            product.Rows.Add(3, "Giáo trình Photoshop", 20, 65000);
            product.Rows.Add(4, "Triết học", 7, 15000);
            product.Rows.Add(5, "Lập trình mạng Cisco", 2, 21000);
            product.Rows.Add(6, "Làm chủ Microsoft Office 2019", 3, 89000);
            product.Rows.Add(7, "Lập trình hướng đối tượng JAVA", 1, 150000);
            product.Rows.Add(8, "Giáo trình Android/IOS", 8, 90000);
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
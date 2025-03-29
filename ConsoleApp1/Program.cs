using System;
using System.Data;
using System.IO;
using ExcelDataReader;

class Program
{
    static void Main()
    {
        string filePath = @"E:\Huflit\BDTKPM\TestCase_NguyenThien.xlsx";
        
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        
        using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        using (var reader = ExcelReaderFactory.CreateReader(stream))
        {
            var result = reader.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = true  // Sử dụng dòng đầu tiên làm tiêu đề cột
                }
            });

            // Kiểm tra danh sách sheet
            Console.WriteLine("📌 Danh sách sheet trong file:");
            foreach (DataTable table in result.Tables)
            {
                Console.WriteLine("- " + table.TableName);
            }

            // Lấy sheet "TestCase_NguyenThien"
            var table = result.Tables["TestCase_NguyenThien"];
            if (table == null)
            {
                Console.WriteLine("⚠ Lỗi: Không tìm thấy sheet 'TestCase_NguyenThien' trong file Excel.");
                return;
            }

            // In danh sách cột
            Console.WriteLine("\n📌 Danh sách cột:");
            foreach (DataColumn column in table.Columns)
            {
                Console.WriteLine("- " + column.ColumnName);
            }

            // In nội dung của file Excel
            Console.WriteLine("\n📌 Dữ liệu trong file:");
            foreach (DataRow row in table.Rows)
            {
                foreach (var item in row.ItemArray)
                {
                    Console.Write(item + "\t");
                }
                Console.WriteLine();
            }
        }
    }
}

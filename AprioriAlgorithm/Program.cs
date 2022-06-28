using OfficeOpenXml;
using System;
using System.IO;

namespace AprioriAlgorithm
{
    class Program
    {
        static void Main(string[] args)
        {
            FileInfo existingFile = new FileInfo("C:\\Users\\AybukeONAY\\OneDrive\\Masaüstü\\Donem Projesi\\Büyük Veri\\Groceries_dataset.xlsx");
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                //get the first worksheet in the workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                int colCount = worksheet.Dimension.End.Column; //get Column Count
                int rowCount = worksheet.Dimension.End.Row; //get row count
                for (int row = 1; row <= rowCount; row++)
                {
                    for (int col = 1; col <= colCount; col++)
                    {
                        Console.WriteLine(" Row:" + row + " column:" + col + " Value:" + worksheet.Cells[row, col].Value.ToString().Trim());
                    }
                }
            }
        }
    }
}

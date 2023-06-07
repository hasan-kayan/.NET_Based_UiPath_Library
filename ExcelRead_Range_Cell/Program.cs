using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelReadApp
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Enter the range to read (e.g., A1:B5):");
            string range = Console.ReadLine();

            Console.WriteLine("Enter the path to the Excel file:");
            string filePath = Console.ReadLine();

            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);
            Excel.Worksheet worksheet = workbook.ActiveSheet;

            try
            {
                Excel.Range excelRange = worksheet.Range[range];
                object[,] values = excelRange.Value;

                int rowCount = values.GetLength(0);
                int columnCount = values.GetLength(1);

                Console.WriteLine($"Reading range: {range}");
                Console.WriteLine();

                for (int row = 1; row <= rowCount; row++)
                {
                    for (int column = 1; column <= columnCount; column++)
                    {
                        object value = values[row, column];
                        Console.Write(value + "\t");
                    }
                    Console.WriteLine();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
            finally
            {
                workbook.Close();
                excelApp.Quit();
            }

            Console.ReadLine();
        }
    }
}

using System;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Text;

namespace ExcelReadApp
{
    class Program
    {
        static void Main(string[] args)
        {
            // Set the console output encoding to UTF-8
            Console.OutputEncoding = Encoding.UTF8;

            Console.WriteLine("Enter the path to the Excel file:");
            string filePath = Console.ReadLine();

            Console.WriteLine("Enter the range to read (e.g., A1:B5):");
            string range = Console.ReadLine();

            Console.WriteLine("Please enter the worksheet name you want to work on:");
            string worksheetName = Console.ReadLine();

            Application app = new Application();
            app.Visible = false;

            Workbook existingWorkbook = null;
            Worksheet worksheet = null;

            try
            {
                // Convert the file path string to UTF-8 encoding
                byte[] filePathBytes = Encoding.UTF8.GetBytes(filePath);
                string encodedFilePath = Encoding.UTF8.GetString(filePathBytes);

                existingWorkbook = app.Workbooks.Open(encodedFilePath); // Open file to read
                worksheet = existingWorkbook.Worksheets[worksheetName]; // Declare Worksheet

                Range excelRange;
                if (string.IsNullOrEmpty(range))
                {
                    excelRange = worksheet.UsedRange;
                }
                else
                {
                    excelRange = worksheet.Range[range];
                }

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
                if (worksheet != null)
                {
                    Marshal.ReleaseComObject(worksheet);
                }
                if (existingWorkbook != null)
                {
                    existingWorkbook.Close();
                    Marshal.ReleaseComObject(existingWorkbook);
                }
                if (app != null)
                {
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                }

                worksheet = null;
                existingWorkbook = null;
                app = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            Console.ReadLine();
        }
    }
}


using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using static System.Net.Mime.MediaTypeNames;

namespace ExcelDataWriter
{
    class Program
    {
        static void Main(string[] args)
        {
            // Set Turkish culture for proper character handling
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("tr-TR");

            // Prompt the user to enter the worksheet name
            Console.WriteLine("Enter the name of the worksheet:");
            string worksheetName = Console.ReadLine();

            // Prompt the user to enter the range where the data will be written
            Console.WriteLine("Enter the range (e.g., A1:B3) where the data will be written:");
            string range = Console.ReadLine();

            // Prompt the user to enter the data to be written
            Console.WriteLine("Enter the data to be written (comma-separated values for each row):");
            string dataInput = Console.ReadLine();

            // Split the data into rows
            string[] dataRows = dataInput.Split(',');

            // Get the currently active Excel application
            Application excelApp = (Application)Marshal.GetActiveObject("Excel.Application");

            // Get the active workbook
            Workbook workbook = excelApp.ActiveWorkbook;

            // Get the worksheet by name
            Worksheet worksheet = (Worksheet)workbook.Sheets[worksheetName];

            // Get the range to write the data
            Range writeRange = worksheet.Range[range];

            // Convert the data into a 2D array
            int rowCount = dataRows.Length;
            int columnCount = range.Split(':').Length / 2;
            object[,] data = new object[rowCount, columnCount];

            for (int row = 0; row < rowCount; row++)
            {
                string[] rowData = dataRows[row].Split(',');
                for (int col = 0; col < columnCount; col++)
                {
                    data[row, col] = rowData[col];
                }
            }

            // Write the data to the range
            writeRange.Value = data;

            // Clean up Excel objects
            Marshal.ReleaseComObject(writeRange);
            Marshal.ReleaseComObject(worksheet);
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(excelApp);

            Console.WriteLine("Data written successfully.");

            Console.WriteLine("Press any key to exit.");
            Console.ReadKey();
        }
    }
}

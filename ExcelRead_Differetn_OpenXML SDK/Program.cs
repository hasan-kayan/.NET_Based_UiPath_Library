using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using static System.Net.Mime.MediaTypeNames;

namespace ExcelDataReader
{
    class Program
    {
        static void Main(string[] args)
        {
            // Prompt the user to enter the range to read from Excel
            Console.WriteLine("Enter the range to read (e.g., A1:B5):");
            string range = Console.ReadLine();

            // Create an Excel application object
            Application excelApp = null;
            try
            {
                excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
            }
            catch (COMException)
            {
                Console.WriteLine("No active Excel instance found.");
                return;
            }

            // Get the active workbook
            Workbook workbook = excelApp.ActiveWorkbook;
            if (workbook == null)
            {
                Console.WriteLine("No open workbook found.");
                return;
            }

            // Get the active worksheet
            Worksheet worksheet = workbook.ActiveSheet;

            // Read the data from the specified range
            Range excelRange = worksheet.Range[range];
            object[,] data = excelRange.Value;

            // Display the data in the console
            int rowCount = data.GetLength(0);
            int columnCount = data.GetLength(1);

            for (int row = 1; row <= rowCount; row++)
            {
                for (int col = 1; col <= columnCount; col++)
                {
                    object cellValue = data[row, col];
                    Console.Write(cellValue + "\t");
                }
                Console.WriteLine();
            }

            // Clean up Excel objects
            Marshal.ReleaseComObject(excelRange);
            Marshal.ReleaseComObject(worksheet);
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(excelApp);

            Console.WriteLine("Press any key to exit.");
            Console.ReadKey();
        }
    }
}

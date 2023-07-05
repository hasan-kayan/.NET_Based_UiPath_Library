using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace ExcelDataReader
{
    class Program
    {
        static void Main(string[] args)
        {
            // Prompt the user to enter the Excel file path
            Console.WriteLine("Enter the Excel file path:");
            string filePath = Console.ReadLine();

            // Create an Excel application object
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            try
            {
                excelApp = new Microsoft.Office.Interop.Excel.Application();
            }
            catch (COMException)
            {
                Console.WriteLine("Failed to create Excel application object.");
                return;
            }

            // Open the workbook
            Workbook workbook = excelApp.Workbooks.Open(filePath);
            if (workbook == null)
            {
                Console.WriteLine("Failed to open the workbook.");
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
                return;
            }

            // Get the first worksheet in the workbook
            Worksheet worksheet = workbook.Sheets[1];

            // Prompt the user to enter the range to read from Excel
            Console.WriteLine("Enter the range to read (e.g., A1:B5):");
            string range = Console.ReadLine();

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
            workbook.Close();
            excelApp.Quit();
            Marshal.ReleaseComObject(excelRange);
            Marshal.ReleaseComObject(worksheet);
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(excelApp);

            Console.WriteLine("Press any key to exit.");
            Console.ReadKey();
        }
    }
}

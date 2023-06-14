using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using System.IO;

namespace ExcelDataReader
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create an Excel application object
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            try
            {
                excelApp = (Microsoft.Office.Interop.Excel.Application)Marshal.GetActiveObject("Excel.Application");
            }
            catch (COMException)
            {
                Console.WriteLine("No active Excel instance found.");
                return;
            }

            Console.WriteLine("Excel Detected");

            // Get the active workbook
            Workbook workbook = excelApp.ActiveWorkbook;
            if (workbook == null)
            {
                Console.WriteLine("No open workbook found.");
                return;
            }
            Console.WriteLine("Workbook Detected");
            Console.WriteLine(workbook.Name);

            // Get the active worksheet
            Worksheet worksheet = workbook.ActiveSheet;
            Console.WriteLine("Worksheet Detected '" + worksheet.Name + "'");

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

            // Write the data array into a file
            string filePath = "output.txt";
            using (StreamWriter writer = File.CreateText(filePath))
            {
                for (int row = 1; row <= rowCount; row++)
                {
                    for (int col = 1; col <= columnCount; col++)
                    {
                        object cellValue = data[row, col];
                        writer.Write(cellValue + "\t");
                    }
                    writer.WriteLine();
                }
            }

            Console.WriteLine("Data array written to file: " + filePath);

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

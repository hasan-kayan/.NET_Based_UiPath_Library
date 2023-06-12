using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using static System.Net.Mime.MediaTypeNames;

namespace TurkishExcelReader
{
    class Program
    {
        static void Main(string[] args)
        {
            // Set Turkish culture for proper character handling
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("tr-TR");

            // Prompt the user to enter the file path
            Console.WriteLine("Enter the file path of the Excel file:");
            string filePath = Console.ReadLine();

            // Prompt the user to enter the worksheet name
            Console.WriteLine("Enter the name of the worksheet:");
            string worksheetName = Console.ReadLine();

            // Create an Excel application object
            Application excelApp = null;
            try
            {
                excelApp = new Application();
            }
            catch (COMException)
            {
                Console.WriteLine("Failed to create Excel application.");
                return;
            }

            // Open the workbook
            Workbook workbook = null;
            try
            {
                workbook = excelApp.Workbooks.Open(filePath);
            }
            catch (COMException)
            {
                Console.WriteLine("Failed to open the Excel file.");
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
                return;
            }

            // Get the worksheet by name
            Worksheet worksheet = null;
            try
            {
                worksheet = (Worksheet)workbook.Sheets[worksheetName];
            }
            catch (COMException)
            {
                Console.WriteLine("Failed to find the worksheet.");
                workbook.Close();
                excelApp.Quit();
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(excelApp);
                return;
            }

            // Read the data from the worksheet
            Range usedRange = worksheet.UsedRange;
            object[,] data = usedRange.Value;

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
            Marshal.ReleaseComObject(usedRange);
            Marshal.ReleaseComObject(worksheet);
            workbook.Close();
            Marshal.ReleaseComObject(workbook);
            excelApp.Quit();
            Marshal.ReleaseComObject(excelApp);

            Console.WriteLine("Press any key to exit.");
            Console.ReadKey();
        }
    }
}


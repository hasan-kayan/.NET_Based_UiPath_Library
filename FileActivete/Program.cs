using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace TurkishExcelReader
{
    class Program
    {
        static void Main(string[] args)
        {
            // Set Turkish culture for proper character handling
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("tr-TR");

            // Prompt the user to enter the file path
            // Console.WriteLine("Enter the file path of the Excel file:");
            // string filePath = Console.ReadLine();

            // Prompt the user to enter the worksheet name
            //Console.WriteLine("Enter the name of the worksheet:");
            //string worksheetName = Console.ReadLine();

            string worksheetName = "Sayfa1";
            string filePath = "C:\\Users\\hasan\\Desktop\\İşBankasıDeneme.xlsx";



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

            Console.WriteLine("Excel Application Created");


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
            Console.WriteLine("Opened Workbook");


            // Get the worksheet by name
            Worksheet worksheet = null;
            try
            {
                worksheet = (Worksheet)workbook.Sheets[worksheetName];
            }
            catch (COMException)
            {
                Console.WriteLine("Failed to find the worksheet.");
                
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(excelApp);
                return;
            }

         
        }
    }
}

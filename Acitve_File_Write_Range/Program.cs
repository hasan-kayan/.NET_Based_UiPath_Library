using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

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

            // Prompt the user to enter the column name where the data will be written
            Console.WriteLine("Enter the column name (e.g., A) where the data will be written:");
            string columnName = Console.ReadLine();

            // Prompt the user to enter the data to be written
            Console.WriteLine("Enter the data to be written, separete data by comma '|':");
            String data = Console.ReadLine();

            string[] Data = data.Split('|'); // One line string will be seperated | bunula ayır 



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

            // Get the active worksheet
            Worksheet worksheet = workbook.ActiveSheet;

        }
    }
}

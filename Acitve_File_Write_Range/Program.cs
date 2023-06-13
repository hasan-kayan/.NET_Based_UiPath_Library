using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace ExcelDataReader
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Enter the column name you want to add the data:");
            string columnName = Console.ReadLine();

            Console.WriteLine("Enter data you want to add, separated by '|':");
            string input = Console.ReadLine();

            string[] data = input.Split('|');
            Console.WriteLine(data);

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

            Console.WriteLine("Excel detected.");

            // Get the active workbook
            Workbook workbook = null;
            try
            {
                workbook = excelApp.ActiveWorkbook;
            }
            catch (COMException)
            {
                Console.WriteLine("No open workbook found.");
                Marshal.ReleaseComObject(excelApp);
                return;
            }

            Console.WriteLine("Workbook detected.");

            Console.WriteLine(workbook.Name);



            // Get the active worksheet
            Worksheet worksheet = workbook.ActiveSheet;

            for (int i = 0; i < data.Length; i++)
            {
                worksheet.Range[columnName + (2 + i)].Value = data[i];
            }

            workbook.Save();

            Console.WriteLine("Press any key to exit.");
            Console.ReadKey();
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Excel;

namespace ExcelApp_Write_Range
{
    internal class Program
    {
        static void Main(string[] args)
        {

            Console.WriteLine("Enter the path to the Excel file:");
            string filePath = Console.ReadLine();

            Console.WriteLine("Enter the column to write (e.g., A):");
            string range = Console.ReadLine();

            Console.WriteLine("Please enter the worksheet name you want to work on:");
            string worksheetName = Console.ReadLine();

            Console.Write("Enter the values separated by commas: "); // Take input data from user 
            string input = Console.ReadLine();

            string[] Data = input.Split(','); // One line string will be seperated

            Application app = new Application();
            app.Visible = true;

            Workbook existingWorkbook = app.Workbooks.Open(filePath); // Open file to read
            Worksheet worksheet = existingWorkbook.Worksheets[worksheetName];

            

            double[] SalesDate = { 4.3, 4, 21, 324, 17 };

            for (int i = 0; i < Data.Length; i++)
            {
                worksheet.Range[range + (2 + i)].Value = Data[i];
            }

            existingWorkbook.Save();
            existingWorkbook.Close();
            app.Quit();

           
        }
    }
}
// C:\Users\hasan\Desktop\excel applications try\deneme.xlsx
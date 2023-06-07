using System;
using System.Threading.Tasks;
using static System.Net.Mime.MediaTypeNames;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace ExcelReadApp
{
    class Program
    {
        static void Main(string[] args)
        {

            Console.WriteLine("Enter the path to the Excel file:");
            string filePath = Console.ReadLine();

            Console.WriteLine("Enter the range to read (e.g., A1:B5):");
            string range = Console.ReadLine();

            Console.WriteLine("Please enter the worksheet name you want to work on:");
            string worksheetName = Console.ReadLine();
            

            Application app = new Application();
            app.Visible = true;

            Workbook existingWorkbook = app.Workbooks.Open(filePath); // Open file to read
            Worksheet worksheet = existingWorkbook.Worksheets[worksheetName]; // Declare Worksheet



            try // For possibel mistakes, try-catch 
            {
                Range excelRange = worksheet.Range[range]; // Range structre is embedded into lib 
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
                existingWorkbook.Close();
                app.Quit();
            }

            Console.ReadLine();
        }
    }
}

// C:\Users\hasan\Desktop\excel applications try\deneme.xlsx
// BIG DATA //
// C:\Users\hasan\Desktop\excel applications try\Halkbank\TC Hazine ve Maliye Bakanlığı yazısı - İhracat bedelleri+IBKB_V2_Exa (YENİ)_995_03.30.2023_11.50.47.xlsx
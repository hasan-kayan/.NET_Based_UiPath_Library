using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.IO;

using DataTable = System.Data.DataTable;

using Excel = Microsoft.Office.Interop.Excel;
using System.Linq.Expressions;
using System.Data;

namespace All_big_data_reading
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string excelFilePath = @"";

            string sheetName = "";

            string range = "";



            // Excel uygulamasını başlat
            Console.WriteLine("New Application Starting...");
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = false;

            // Aktif çalışma kitabını ve çalışma sayfasını al
            Console.WriteLine("Workbook detection...");
            Excel.Workbook workbook = excelApp.Workbooks.Open(excelFilePath);
            Excel.Worksheet worksheet = workbook.Sheets[sheetName];


            Excel.Worksheet newWorksheet = null;

            try
            {
                Console.WriteLine("Copying Cells");
                // Tüm hücreleri seç ve kopyala
                Excel.Range cells = worksheet.Cells;
                cells.Select();
                cells.Copy();

                Console.WriteLine("Creating New Sheet");

                // Yeni çalışma sayfasını ekleyin
                newWorksheet = workbook.Sheets.Add(After: workbook.ActiveSheet);

                // Yeni sayfanın adını değiştirin
                newWorksheet.Name = "Copied";
                Console.WriteLine("New Sheet named 'Copied' Created");

                // Yapıştırma işlemini gerçekleştir
                Excel.Range pasteRange = newWorksheet.Cells;
                pasteRange.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone,
                    SkipBlanks: false, Transpose: false);

                Console.WriteLine("Page Copied Successfully!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error occurred: {ex.Message}");
            }


            // Read the data
            Excel.Range excelRange = null;

            try
            {
                // Read data from the new worksheet
                if (string.IsNullOrWhiteSpace(range))
                {
                    excelRange = newWorksheet.UsedRange;
                }
                else
                {
                    excelRange = newWorksheet.Range[range];
                }

                // Continue with reading the data and creating DataTable
                object[,] excelData = (object[,])excelRange.Value;

                // Create a DataTable and transfer the data
                DataTable dataTable = new DataTable();
                int rowCount = excelData.GetLength(0);
                int columnCount = excelData.GetLength(1);

                for (int col = 1; col <= columnCount; col++)
                {
                    string columnName = excelData[1, col]?.ToString() ?? $"Column{col}";
                    dataTable.Columns.Add(columnName);
                }

                for (int row = 2; row <= rowCount; row++)
                {
                    DataRow dataRow = dataTable.NewRow();
                    for (int col = 1; col <= columnCount; col++)
                    {
                        dataRow[col - 1] = excelData[row, col];
                    }
                    dataTable.Rows.Add(dataRow);
                }

                // Use the DataTable as needed
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error occurred: {ex.Message}");
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);






















        }
    }
    }


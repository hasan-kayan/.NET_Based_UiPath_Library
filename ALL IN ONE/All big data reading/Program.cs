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
            string excelFilePath = @"C:\Users\hasan\Desktop\Büyük Excel\Halkbank\TC Hazine ve Maliye Bakanlığı yazısı - İhracat bedelleri+IBKB_V2_Exa_Robotik Süreç (1).xlsx";
            string sheetName = "TC Hazine ve Maliye Bakanlığı y";
            string range = "B4:BJ120000";
           
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

                // Yeni sayfayı aktif hale getirin
                newWorksheet.Activate();

             
                Excel.Range excelRange = null;

                // Yeni çalışma sayfasının null olup olmadığını kontrol edin
                if (newWorksheet != null)
                {
                    // Read data from the new worksheet
                    if (string.IsNullOrWhiteSpace(range))
                    {
                        Console.WriteLine("All data reading");
                        excelRange = newWorksheet.UsedRange;
                    }
                    else
                    {
                        Console.WriteLine("The range : " + range + "is reading");
                        excelRange = newWorksheet.Range[range];
                    }

                    // Continue with reading the data and creating DataTable
                    Console.WriteLine("Data Table Creating, data reading...");
                    object[,] excelData = (object[,])excelRange.Value;

                    // Create a DataTable and transfer the data
                    DataTable dataTable = new DataTable();
                    int rowCount = excelData.GetLength(0);
                    int columnCount = excelData.GetLength(1);

                    Console.WriteLine("Data Table created Succesfully");

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
                    Console.WriteLine("Data Table Content:");
                    foreach (DataRow row in dataTable.Rows)
                    {
                        foreach (DataColumn col in dataTable.Columns)
                        {
                            Console.Write(row[col] + "\t");
                        }
                        Console.WriteLine();
                    }
                }
                else
                {
                    Console.WriteLine("New worksheet is null. Cannot proceed with data reading.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error occurred: {ex.Message}");
            }

            // Serbest bırakma işlemleri
            
         
            workbook?.Close();
            excelApp?.Quit();

           
            System.Runtime.InteropServices.Marshal.ReleaseComObject(newWorksheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

            GC.Collect();
            GC.WaitForPendingFinalizers();
            Console.ReadKey();
        }
    }
}

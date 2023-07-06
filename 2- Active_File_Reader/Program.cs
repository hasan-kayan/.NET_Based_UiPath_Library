using System;
using Microsoft.Office.Interop.Excel;
using System.Data;
using DataTable = System.Data.DataTable;

class Program
{
    static void Main()
    {
        string excelFilePath = @"C:\Users\hasan\Desktop\Büyük Excel\Halkbank\TC Hazine ve Maliye Bakanlığı yazısı - İhracat bedelleri+IBKB_V2_Exa (YENİ)_995_03.30.2023_11.50.47.xlsx";
        string sheetName = "TC Hazine ve Maliye Bakanlığı y";
        string range = "B4:BJ200000"; // Örnek aralık
        int rowsPerIteration = 4000;

        try
        {
            Console.WriteLine("Excel dosyası açılıyor...");
            Application excelApp = new Application();
            Workbook workbook = null;
            Worksheet worksheet = null;

            // Excel dosyasının zaten açık olup olmadığını kontrol et
            foreach (Workbook openWorkbook in excelApp.Workbooks)
            {
                if (openWorkbook.FullName == excelFilePath)
                {
                    workbook = openWorkbook;
                    worksheet = workbook.Sheets[sheetName];
                    break;
                }
            }

            if (workbook == null)
            {
                Console.WriteLine("Excel dosyası zaten açık değil. Dosya açılıyor...");
                workbook = excelApp.Workbooks.Open(excelFilePath);
                worksheet = workbook.Sheets[sheetName];
            }

            if (worksheet != null)
            {
                Console.WriteLine("Sayfa bulundu. Hücre değerleri alınıyor...");
                Range cells = worksheet.Range[range];
                int rowCount = cells.Rows.Count;
                int currentRow = 1;

                while (currentRow <= rowCount)
                {
                    int endRow = Math.Min(currentRow + rowsPerIteration - 1, rowCount);
                    string currentRange = $"B{currentRow}:BJ{endRow}";
                    Range currentCells = worksheet.Range[currentRange];
                    object[,] cellValues = (object[,])currentCells.Value;

                    if (cellValues != null)
                    {
                        int columnCount = cellValues.GetLength(1);

                        DataTable dataTable = new DataTable();

                        // DataTable'a sütunları ekle
                        for (int col = 1; col <= columnCount; col++)
                        {
                            dataTable.Columns.Add($"Column{col}");
                        }

                        // Hücre değerlerini DataTable'a ekle
                        for (int row = 1; row <= rowsPerIteration; row++)
                        {
                            if (currentRow > rowCount)
                                break;

                            DataRow dataRow = dataTable.NewRow();

                            for (int col = 1; col <= columnCount; col++)
                            {
                                object cellValue = cellValues[row, col];
                                dataRow[col - 1] = cellValue;
                            }

                            dataTable.Rows.Add(dataRow);
                            currentRow++;
                        }

                        Console.WriteLine("Hücre değerleri alındı. DataTable çıktısı:");
                        PrintDataTable(dataTable);
                    }
                    else
                    {
                        Console.WriteLine("Belirtilen aralıkta hücre değeri bulunamadı.");
                    }
                }
            }
            else
            {
                Console.WriteLine("Belirtilen sayfa adı bulunamadı.");
            }

            workbook.Close();
            excelApp.Quit();
        }
        catch (Exception ex)
        {
            Console.WriteLine("Bir hata oluştu: " + ex.Message);
        }

        Console.ReadLine();
    }

    static void PrintDataTable(DataTable dataTable) 
    {
        foreach (DataColumn column in dataTable.Columns)
        {
            Console.Write($"{column.ColumnName}\t");
        }
        Console.WriteLine();

        foreach (DataRow row in dataTable.Rows)
        {
            foreach (var item in row.ItemArray)
            {
                Console.Write($"{item}\t");
            }
            Console.WriteLine();
        }
    }
}

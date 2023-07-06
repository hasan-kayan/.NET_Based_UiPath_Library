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

        string startColumnName = "A";
        int startColumnIndex = 1;

        string endColumnName = "BJ";
        int endColumnIndex = 20000;

        int batchSize = 4000;

        try
        {
            Console.WriteLine("Opening the Excel file...");
            Application excelApp = new Application();
            Workbook workbook = null;
            Worksheet worksheet = null;

            // Check if the Excel file is already open
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
                Console.WriteLine("Excel file is not already open. Opening the file...");
                workbook = excelApp.Workbooks.Open(excelFilePath);
                worksheet = workbook.Sheets[sheetName];
            }

            if (worksheet != null)
            {
                Console.WriteLine("Sheet found. Retrieving cell values...");

                int rowCount = worksheet.Cells.Rows.Count;
                int columnCount = worksheet.Cells.Columns.Count;

                DataTable dataTable = new DataTable();

                // Add columns to the DataTable
                for (int col = 1; col <= columnCount; col++)
                {
                    dataTable.Columns.Add($"Column{col}");
                }

                // Read data in batches
                for (int startRow = 1; startRow <= rowCount; startRow += batchSize)
                {
                    int endRow = Math.Min(startRow + batchSize - 1, rowCount);

                    string range = $"{startColumnName}{startRow}:{endColumnName}{endRow}";
                    Range cells = worksheet.Range[range];
                    object[,] cellValues = (object[,])cells.Value;

                    if (cellValues != null)
                    {
                        // Add cell values to the DataTable
                        for (int row = 1; row <= cellValues.GetLength(0); row++)
                        {
                            DataRow dataRow = dataTable.NewRow();

                            for (int col = 1; col <= cellValues.GetLength(1); col++)
                            {
                                object cellValue = cellValues[row, col];
                                dataRow[col - 1] = cellValue;
                            }

                            dataTable.Rows.Add(dataRow);
                        }
                    }
                }

                Console.WriteLine("Cell values retrieved. DataTable output:");
                PrintDataTable(dataTable);
            }
            else
            {
                Console.WriteLine("Specified sheet name not found.");
            }

            workbook.Close();
            excelApp.Quit();
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
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

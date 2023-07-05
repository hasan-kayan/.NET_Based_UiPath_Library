using System;
using System.Data;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;

class Program
{
    static void Main()
    {
        string excelFilePath = @"C:\Users\hasan\Desktop\Büyük Excel\Halkbank\TC Hazine ve Maliye Bakanlığı yazısı - İhracat bedelleri+IBKB_V2_Exa (YENİ)_995_03.30.2023_11.50.47.xlsx";
        string sheetName = "TC Hazine ve Maliye Bakanlığı y";
        string range = "A1:C10"; // Örnek aralık

        try
        {
            Console.WriteLine("Opening the Excel file...");
            Application excelApp = new Application();
            Workbook workbook = excelApp.Workbooks.Open(excelFilePath);
            Worksheet worksheet = workbook.Sheets[sheetName];

            if (worksheet != null)
            {
                Console.WriteLine("Sheet found. Retrieving cell values...");
                Range cells = worksheet.Range[range];
                object[,] cellValues = (object[,])cells.Value;

                if (cellValues != null)
                {
                    int rowCount = cellValues.GetLength(0);
                    int columnCount = cellValues.GetLength(1);

                    DataTable dataTable = new DataTable();

                    // Add columns to the DataTable
                    for (int col = 1; col <= columnCount; col++)
                    {
                        dataTable.Columns.Add($"Column{col}");
                    }

                    // Add cell values to the DataTable
                    for (int row = 1; row <= rowCount; row++)
                    {
                        DataRow dataRow = dataTable.NewRow();

                        for (int col = 1; col <= columnCount; col++)
                        {
                            object cellValue = cellValues[row, col];
                            dataRow[col - 1] = cellValue;
                        }

                        dataTable.Rows.Add(dataRow);
                    }

                    Console.WriteLine("Cell values retrieved. DataTable output:");
                    PrintDataTable(dataTable);
                }
                else
                {
                    Console.WriteLine("No cell values found in the specified range.");
                }
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

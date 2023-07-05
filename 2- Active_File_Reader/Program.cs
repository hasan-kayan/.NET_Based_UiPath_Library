using System;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;

namespace ExcelHelper
{
    public class ExcelOperations
    {
        public DataTable CopyAndReadWorksheet(string filePath, string sourceSheetName)
        {
            // Create an Excel application object
            Application excelApp = null;
            Workbook workbook = null;

            try
            {
                excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                workbook = excelApp.Workbooks[Path.GetFileName(filePath)];
            }
            catch (Exception)
            {
                excelApp = new Application();
                workbook = excelApp.Workbooks.Open(filePath);
            }

            if (workbook == null)
            {
                excelApp.Quit();
                throw new Exception("Failed to open the workbook.");
            }

            // Get the source worksheet
            Worksheet sourceWorksheet = null;
            foreach (Worksheet worksheet in workbook.Sheets)
            {
                if (worksheet.Name == sourceSheetName)
                {
                    sourceWorksheet = worksheet;
                    break;
                }
            }
            if (sourceWorksheet == null)
            {
                workbook.Close();
                excelApp.Quit();
                throw new Exception($"Worksheet '{sourceSheetName}' not found.");
            }

            // Get the used range in the source worksheet
            Range usedRange = sourceWorksheet.UsedRange;

            // Create a new worksheet
            Worksheet newWorksheet = workbook.Sheets.Add(Type.Missing, workbook.Sheets[workbook.Sheets.Count], Type.Missing, Type.Missing);

            // Copy the used range from source worksheet to new worksheet
            usedRange.Copy(newWorksheet.Cells[1, 1]);

            // Paste values only in the new worksheet
            newWorksheet.PasteSpecial(XlPasteType.xlPasteValues, XlPasteSpecialOperation.xlPasteSpecialOperationNone);

            // Get the data from the new worksheet
            Range newWorksheetUsedRange = newWorksheet.UsedRange;
            object[,] excelData = (object[,])newWorksheetUsedRange.Value;

            // Convert the data to a DataTable
            DataTable dataTable = new DataTable();
            int rowCount = excelData.GetLength(0);
            int columnCount = excelData.GetLength(1);

            // Add columns to the DataTable
            for (int col = 1; col <= columnCount; col++)
            {
                string columnName = excelData[1, col]?.ToString() ?? $"Column{col}";
                dataTable.Columns.Add(columnName);
            }

            // Add rows to the DataTable
            for (int row = 2; row <= rowCount; row++)
            {
                DataRow dataRow = dataTable.NewRow();
                for (int col = 1; col <= columnCount; col++)
                {
                    dataRow[col - 1] = excelData[row, col];
                }
                dataTable.Rows.Add(dataRow);
            }

            // Close the workbook and release Excel objects
            workbook.Close();
            excelApp.Quit();
            ReleaseObject(newWorksheetUsedRange);
            ReleaseObject(newWorksheet);
            ReleaseObject(usedRange);
            ReleaseObject(sourceWorksheet);
            ReleaseObject(workbook);
            ReleaseObject(excelApp);

            return dataTable;
        }

        private void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                throw ex;
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}

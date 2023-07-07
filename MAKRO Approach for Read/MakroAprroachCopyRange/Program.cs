﻿using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

class Program
{
    static void Main()
    {
        Console.Write("Excel dosyasının yolunu girin: ");
        string excelFilePath = Console.ReadLine();

        Console.Write("Sayfa adını girin: ");
        string sheetName = Console.ReadLine();

        // Excel uygulamasını başlat
        Excel.Application excelApp = new Excel.Application();
        excelApp.Visible = true;

        // Aktif çalışma kitabını ve çalışma sayfasını al
        Excel.Workbook workbook = excelApp.Workbooks.Open(excelFilePath);
        Excel.Worksheet worksheet = workbook.Sheets[sheetName];

        try
        {
            // Tüm hücreleri seç ve kopyala
            Excel.Range cells = worksheet.Cells;
            cells.Select();
            cells.Copy();

            // Yeni bir çalışma sayfası ekle
            Excel.Worksheet newWorksheet = workbook.Sheets.Add(After: workbook.ActiveSheet);

            // Yapıştırma işlemini gerçekleştir
            Excel.Range pasteRange = newWorksheet.Cells;
            pasteRange.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone,
                SkipBlanks: false, Transpose: false);

            Console.WriteLine("Kopyalama ve yapıştırma işlemi tamamlandı.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Hata: " + ex.Message);
        }
        finally
        {
            // Excel uygulamasını kapat ve kaynakları serbest bırak
            workbook.Close();
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
        }

        Console.ReadKey();
    }
}

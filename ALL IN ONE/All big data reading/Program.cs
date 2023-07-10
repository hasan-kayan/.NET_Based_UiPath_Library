using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace All_big_data_reading
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Excel uygulamasını başlat
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = true;

            // Aktif çalışma kitabını ve çalışma sayfasını al
            Excel.Workbook workbook = excelApp.Workbooks.Open(excelFilePath);
            Excel.Worksheet worksheet = workbook.Sheets[sheetName];
        }
    }
}

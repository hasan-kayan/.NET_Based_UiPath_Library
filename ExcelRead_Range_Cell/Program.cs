using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Net.Mime.MediaTypeNames;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace ExcelRead_Range_Cell
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Please Enter The File Path:");
            string filePath = Console.ReadLine();

            Console.WriteLine("1)- Single Data , 2)- Data set"); ;

            int caseData = Int32.Parse(Console.ReadLine());
            ;

            switch (caseData)
            {
                case 1:
                    Console.WriteLine("Please Enter Data and Cell You want to write:");
                    string cellData = Console.ReadLine();
                    string cellID = Console.ReadLine();

                    break;

               case 2:
                    Console.WriteLine("Please enter the Data set and range of data you want to enter:");
                    string cellData2 = Console.ReadLine();
                    string range = Console.ReadLine();
                    break;

            }





        }
    }
}


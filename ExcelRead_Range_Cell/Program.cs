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

            Console.WriteLine("1)- Cell Reading , 2)- Range Reading"); ;

            int caseData = Int32.Parse(Console.ReadLine());
            ;

            switch (caseData)
            {
                case 1:
                    Console.WriteLine("Please Enter Data and Cell You want to read:");
                    string cellID = Console.ReadLine();
                    

                    break;

               case 2:
                    Console.WriteLine("Please enter range you want to read:");
                    string range = Console.ReadLine();
                    break;

            }





        }
    }
}


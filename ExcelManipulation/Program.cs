using ExcelCore;
using System;

namespace ExcelManipulation {
    class Program {

       

        static void Main(string[] args) {
            Console.WriteLine("Starting Program");

            Excel excel = new Excel();
            excel.Open("C:\\Users\\royde\\Documents\\Workspace\\AutomationAnywhere\\book1.xlsx");

            Console.WriteLine(excel.GetRows());
        }
    }
}

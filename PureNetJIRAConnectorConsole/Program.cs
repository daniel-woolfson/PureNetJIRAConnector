using System;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace PureNetJIRAConnectorConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            var excelFileHandler = new ExcelFileHandler();
            
            excelFileHandler.createExcelPackage();
            excelFileHandler.SaveFile();
        }

    }
}

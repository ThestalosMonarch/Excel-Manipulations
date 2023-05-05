using System.IO;
using System.Data;
using System.Data.OleDb;
using System;
using Microsoft.Office.Interop.Excel;

namespace Excel
{
    internal class Program
    {
        static void Main(string[] args)
        {

            string strPath = @"C:\Users\lucas\source\repos\Excel\teste.xlsb";
            ExcelManipulations.ConvertFromXLSBToXLSX(strPath);

        }
    }
}

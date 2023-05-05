using Microsoft.Office.Interop.Excel;
using System;
using System.IO;

namespace Excel
{
    public class ExcelManipulations
    {
        public static string ConvertFromXLSBToXLSX(string filepath)
        {
            string strNewPath = "";
            if (!File.Exists(filepath.Replace("xlsb", "xlsx")))
            {
                try
                {
                    Microsoft.Office.Interop.Excel.Application excelApplication = new Microsoft.Office.Interop.Excel.Application();
                    Workbooks workbooks = excelApplication.Workbooks;
                    // open book in any format
                    Workbook workbook = workbooks.Open(filepath, XlUpdateLinks.xlUpdateLinksNever, true, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                    // save in XlFileFormat.xlExcel12 format which is XLSB
                    workbook.SaveAs(filepath.Replace("xlsb", "xlsx"), XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                    // close workbook
                    workbook.Close(false, Type.Missing, Type.Missing);

                    excelApplication.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbooks);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApplication);
                    strNewPath = filepath.Replace("xlsb", "xlsx");
                }
                catch (Exception ex)
                {

                }
                finally
                {
                    //foreach (System.Diagnostics.Process proc in System.Diagnostics.Process.GetProcessesByName("EXCEL"))
                    //{
                    //    proc.Kill();
                    //}
                }
            }
            else
            {
                strNewPath = filepath.Replace("xlsb", "xlsx");
            }
            return strNewPath;

        }
    }
}
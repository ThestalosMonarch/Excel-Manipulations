using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ExcelActions
{
    public class Helpers
    {
        public static string GetExcelColumnName(int columnNumber)
        {
            string columnName = "";

            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                columnNumber = (columnNumber - modulo) / 26;
            }

            return columnName;
        }

        public static int GetExcelColumnPosition(string columnLetter)
        {
            columnLetter = columnLetter.ToUpper();
            // Convert column letter to column index
            int columnPosition = 0;
            foreach (char c in columnLetter)
            {
                columnPosition *= 26;
                columnPosition += (c - 'A' + 1);
            }
            return columnPosition;
        }
        public static int GetLastColumn(Range columnRange)
        {
           int position = columnRange.Cells.Find(What: "*", After: columnRange.Cells[1, columnRange.Columns.Count], LookIn: XlFindLookIn.xlValues, SearchOrder: XlSearchOrder.xlByColumns, SearchDirection: XlSearchDirection.xlPrevious).Column;
           return position;
        }
        public static int GetLastRow(Worksheet worksheet)
        {
            int lastRowNumber = worksheet.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, XlSearchOrder.xlByRows, XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
            return lastRowNumber;
        }

        public static List<string> ColumnToList(Workbook excelWorkbook, string sheetName, int columnIndex)
        {
            List<string> columnData = new List<string>();
            Worksheet excelWorksheet = excelWorkbook.Sheets[sheetName];
            Range excelRange = excelWorksheet.UsedRange;

            int lastRow = GetLastRow(excelWorksheet);
    
            for (int i = 2; i <= lastRow; i++)
            {
                Range cell = excelRange.Cells[i, columnIndex];
                string cellValue = cell.Value2?.ToString(); // Read the cell value
                columnData.Add(cellValue);
            }
            return columnData;
        }

        public static void ChangeColumnName(Worksheet worksheet, int index, string name)
        {
            Microsoft.Office.Interop.Excel.Range column = worksheet.Columns[index];
            column.Cells[1, 1].Value = name;
        }
    }

}

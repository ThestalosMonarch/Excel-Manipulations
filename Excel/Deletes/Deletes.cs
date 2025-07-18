using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;


namespace ExcelActions
{
    public class Deletes
    {
        public static void DeleteHiddenRows(Workbook workbook, string sheetName, int columnIndex)
        {
            Worksheet worksheet = workbook.Worksheets[sheetName];
            Range range = worksheet.UsedRange;

            for (int i = range.Rows.Count; i >= 1; i--)
            {
                Range row = range.Rows[i];
                if (row.Hidden)
                {
                    row.Delete();
                }
            }
        }
        public static void DeleteHiddenRowsUsingRangeCopy(Workbook workbook, string sheetName)
        {
            Worksheet worksheet = workbook.Worksheets[sheetName];
            Range Range = worksheet.UsedRange;

            Range visibleRange = Range.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeVisible);

            // Check if there are visible cells
            if (visibleRange != null)
            {
                // You can copy the visible range to another sheet or range
                Worksheet destinationSheet = workbook.Worksheets.Add();
                visibleRange.Copy(Type.Missing);
                destinationSheet.Range["A1"].PasteSpecial(XlPasteType.xlPasteAllExceptBorders);
                worksheet.Delete();
                destinationSheet.Name = "Base Localidade";
                workbook.Save();
            }
        }

    }
}

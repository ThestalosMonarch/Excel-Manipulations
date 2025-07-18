using Gratificacao_Fronteira;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace ExcelActions
{
    public class Formulas
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="customLogger"></param>
        /// <returns></returns>
        /// 
        public static void InsertFormulas(Workbook templateWorkbook, JObject formulasVlookUp, CustomLogger customLogger)
        {
            customLogger.LogMethodStart();

            Worksheet toInsert = templateWorkbook.Worksheets["Base Localidade"];
            
            Range getLastRows = toInsert.Range["A:A"];
            //FIND the reference of last row and subtract one to not get the footer.
            int lastRow = getLastRows.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, XlSearchOrder.xlByRows, XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
            Range columnRange = null;
  
            foreach (var formula in formulasVlookUp)
            {
                string defaultFormula = formula.Value.ToString(); 
                string columnLetter = formula.Key;
                customLogger.LogMessage("Coluna - " + columnLetter);
                object[,] formulaArray = new object[lastRow - 1, 1];

                columnRange = toInsert.Range[columnLetter + "2:" + columnLetter + lastRow];
                string changedFormula = "";
                for (int i = 1; i < lastRow; i++)
                {
                    changedFormula = ReplacePositionFormula(defaultFormula, (i + 1));
                    formulaArray[i - 1, 0] = changedFormula;
                }
                columnRange.Formula = formulaArray;
            }
            templateWorkbook.Save();
            customLogger.LogMethodEnd();
        }
        public static object[,] GetFormulasFromTemplate(Worksheet worksheet, CustomLogger customLogger)
        {
            customLogger.LogMethodStart();
            int lastRowWithData = worksheet.Range["A:A"].Cells.Find("Formulas - End", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, XlSearchOrder.xlByRows, XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row - 1;

            Range formulaRange = worksheet.Range["A2:B" + lastRowWithData]; // Adjust the range as neededlas

            // Read data from the range into a 2D array
            object[,] formulaData = (object[,])formulaRange.Formula;
            for (int row = 1; row <= formulaData.GetLength(0); row++)
            {
                // Get the formula and column letter from the array
                string formula = formulaData[row, 1]?.ToString(); // Column A
                string columnLetter = formulaData[row, 2]?.ToString(); // Column B
                // Process the formula and column letter as needed
            }
            customLogger.LogMethodEnd();
            return formulaData;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public static string ReplacePositionFormula(string formula, int position)
        {
            if (formula.Contains("NEXTROW"))
            {
                formula = formula.Replace("NEXTROW", (position + 1).ToString());
            }
            if (formula.Contains("LASTROW"))
            {
                formula = formula.Replace("LASTROW", (position - 1).ToString());
            }
            if (formula.Contains("ROW"))
            {
                formula = formula.Replace("ROW", position.ToString());
            }

            return formula;
        }
    }
}

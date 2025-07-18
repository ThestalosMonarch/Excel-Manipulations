using Microsoft.Office.Interop.Excel;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System;
using System.Net.Http;
using System.Windows;
using Gratificacao_Fronteira;
using System.Collections.Generic;

namespace ExcelActions
{
    /// <summary>
    /// Have all methods to copies data that can be made it in a excel file
    /// </summary>
    /// 

    public class Copies
    {
        /// <summary>
        /// Copy Data between two files, considering if the header will be included.
        /// </summary>
        /// <param name="workbookDestination"></param>
        /// <param name="workbookSource"></param>
        /// <param name="customLogger"></param>
        /// <param name="sheetDestination"></param>
        /// <param name="sheetSource"></param>
        /// <param name="excludeHeaders"></param>
        public static void CopyDataBetweenFiles(Workbook workbookDestination, Workbook workbookSource, CustomLogger customLogger,
                                               string sheetDestination,
                                               string sheetSource,
                                               [Optional] bool excludeHeaders,
                                               [Optional] string startPositionSource,
                                               [Optional] string startPositionDestiny) 
        {
            try
            {

                #region If some of the optional parameters will be null, then will add default values 
                customLogger.LogMethodStart();
                
                //sheetDestination = sheetDestination ?? workbookDestination.Sheets[1].Name;
                //sheetSource = sheetSource ?? workbookSource.Sheets[1].Name;

                Worksheet worksheetDestination = workbookDestination.Sheets[sheetDestination];
                Worksheet worksheetSource = workbookSource.Sheets[sheetSource];

                string rangeDestination = "";
                string rangeSource = "";
                
                if (excludeHeaders)
                {
                    int lastRowDestination = Helpers.GetLastRow(worksheetDestination);
                    int lastColumnDestinationNumber = Helpers.GetLastColumn(worksheetDestination.Rows[1]);
                    string lastColumnDestinationName = Helpers.GetExcelColumnName(lastColumnDestinationNumber);


                    int lastRowSource = Helpers.GetLastRow(worksheetSource);
                    int lastColumnSourceNumber = Helpers.GetLastColumn(worksheetSource.Rows[4]);
                    string lastColumnSourceName = Helpers.GetExcelColumnName(lastColumnSourceNumber);
                    int lastRowsSource = Helpers.GetLastRow(worksheetSource);

                    startPositionSource = startPositionSource ?? "A2";
                    startPositionDestiny = startPositionDestiny ?? "A2";

                    rangeDestination = startPositionDestiny +":" + lastColumnDestinationName + lastColumnDestinationNumber;
                    rangeSource = startPositionSource + ":" + lastColumnSourceName + lastRowsSource;
                }
                else
                {
                    rangeDestination = "A1";
                    rangeSource = worksheetSource.UsedRange.Address;
                }
                #endregion

                Range teste = worksheetSource.UsedRange;
                //Clearing destination range before copy
                worksheetDestination.Range[rangeDestination].Clear();

                //Copying the values
                Range destionation = worksheetDestination.Range[rangeDestination];
                Range source = worksheetSource.Range[rangeSource];

                source.Copy(destionation);
                workbookDestination.Save();
            }

            catch(COMException ex)
            {
                customLogger.LogError(ex);
                throw ex;
            }
            catch(Exception ex)
            {
                customLogger.LogError(ex);
                throw ex;
            }
            finally
            {
                customLogger.LogMethodEnd();
            }
        }

        public static void CopieRangeWithoutFormulas(Workbook workbookBase, CustomLogger customLogger)
        {
            

            Worksheet worksheet = workbookBase.Sheets["BASE"]; // Replace with the name of your worksheet

            int lastRow = worksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row;

            List<string> columns = new List<string> { "C", "D", "E", "F", "G", "H", "I", "J", "K", "M", "O", "S", "T", "U", "V", "W" };
            Range range = null;
            foreach (string column in columns)
            {
                range = worksheet.Range[$"{column}14:{column}{lastRow}"];
                object[,] values = range.Value;

                for (int row = 1; row <= values.GetLength(0); row++)
                {
                    values[row, 1] = values[row, 1];
                }

                range.Value = values;
            }
            workbookBase.Save();
            Marshal.ReleaseComObject(range);

        }
    }
}

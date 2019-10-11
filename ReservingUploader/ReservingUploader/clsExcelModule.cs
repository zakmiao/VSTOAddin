using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace ReservingUploader
{
    class clsExcelModule
    {
        //all excel functions

        public void PastToWorksheet(DataTable data)
        {
            //past dataTable to worksheet

            //Write dataTable to array
            var numCols = data.Columns.Count;
            var numRows = data.Rows.Count;
            var myArray = new object[numRows + 2, numCols + 1];
            for (var column = 0; column < numCols; column++)
            {
                myArray[0, column] = data.Columns[column].ColumnName;
                for (var row = 0; row < numRows; row++)
                {
                    myArray[row + 1, column] = data.Rows[row][column];
                }
            }

            //write array to worksheet
            Excel.Application myApp = Globals.ThisAddIn.Application;
            Excel.Workbook myWkbk = myApp.ActiveWorkbook;
            Excel.Worksheet mySheet = myWkbk.ActiveSheet;

            mySheet.Cells.Clear();
            try { mySheet.Range["A1"].Resize[numRows + 1, numCols].Value = myArray; }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        public void PastToWorksheet(DataTable data, Excel.Worksheet destWksht)
        {
            //past dataTable to worksheet

            //Write dataTable to array
            var numCols = data.Columns.Count;
            var numRows = data.Rows.Count;
            var myArray = new object[numRows + 2, numCols + 1];
            for (var column = 0; column < numCols; column++)
            {
                myArray[0, column] = data.Columns[column].ColumnName;
                for (var row = 0; row < numRows; row++)
                {
                    myArray[row + 1, column] = data.Rows[row][column];
                }
            }

            //write array to worksheet
            Excel.Worksheet mySheet = destWksht;
            mySheet.Cells.Clear();
            try { mySheet.Range["A1"].Resize[numRows + 1, numCols].Value = myArray; }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        public void PastToWorksheet(DataTable data, Excel.Range destRange)
        {
            //past dataTable to worksheet

            //Write dataTable to array
            var numCols = data.Columns.Count;
            var numRows = data.Rows.Count;
            var myArray = new object[numRows + 2, numCols + 1];
            for (var column = 0; column < numCols; column++)
            {
                myArray[0, column] = data.Columns[column].ColumnName;
                for (var row = 0; row < numRows; row++)
                {
                    myArray[row + 1, column] = data.Rows[row][column];
                }
            }

            //write array to worksheet
            try { destRange.Resize[numRows + 1, numCols].Value = myArray; }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        public void PastToWorksheetEnd(DataTable data)
        {
            //past dataTable to worksheet

            //Write dataTable to array
            var numCols = data.Columns.Count;
            var numRows = data.Rows.Count;
            var myArray = new object[numRows + 1, numCols + 1];
            for (var column = 0; column < numCols; column++)
            {
                //myArray[0, column] = data.Columns[column].ColumnName;
                for (var row = 0; row < numRows; row++)
                {
                    myArray[row, column] = data.Rows[row][column];
                }
            }

            //write array to worksheet
            Excel.Worksheet mySheet = Globals.ThisAddIn.Application.ThisWorkbook.ActiveSheet;

            //Past Cell

            int pastRow = mySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            Excel.Range PastRange = mySheet.Cells[pastRow, 1];

            Excel.Range myRange = PastRange.Offset[2, 0];

            try { myRange.Resize[numRows, numCols].Value = myArray; }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        public DataTable ExcelToDatatable()
        {
            DataTable tmpXlData = new DataTable();
            Excel.Worksheet mySheet = Globals.ThisAddIn.Application.ThisWorkbook.ActiveSheet;
            Excel.Range myRange = mySheet.UsedRange;

            object[,] XlData = myRange.Value2;

            var nColumn = myRange.Columns.Count;
            var nRow = myRange.Rows.Count;

            for (int column = 1; column <= nColumn; column++)
            {
                if (XlData[1, column]?.ToString() != "")
                {
                    tmpXlData.Columns.Add(XlData[1, column]?.ToString());
                }
            }

            for (int row = 2; row <= nRow; row++)
            {
                DataRow XlDataRow = tmpXlData.NewRow();

                for (int column = 1; column <= nColumn; column++)
                {
                    if (XlData[row, column]?.ToString() != "")
                    {
                        XlDataRow[column - 1] = XlData[row, column]?.ToString();
                    }
                }
                tmpXlData.Rows.Add(XlDataRow);
            }

            return tmpXlData;
        }

        public DataTable ExcelToDatatable(Excel.Workbook myMEWkbk)
        {
            DataTable tmpXlData = new DataTable();
            Excel.Worksheet mySheet = myMEWkbk.Worksheets["ME Flatfile"];
            Excel.Range myRange = mySheet.UsedRange;

            object[,] XlData = myRange.Value2;

            DataColumn ImpRow = new DataColumn("ImportRow");
            ImpRow.DataType = System.Type.GetType("System.Int32");
            ImpRow.AutoIncrement = true;
            ImpRow.AutoIncrementSeed = 1;
            ImpRow.AutoIncrementStep = 1;

            var nColumn = myRange.Columns.Count;
            var nRow = myRange.Rows.Count;

            for (int column = 1; column <= nColumn; column++)
            {
                if (XlData[1, column]?.ToString() != "")
                {
                    tmpXlData.Columns.Add(XlData[1, column]?.ToString());
                }
            }

            tmpXlData.Columns.Add(ImpRow);

            for (int row = 2; row <= nRow; row++)
            {
                DataRow XlDataRow = tmpXlData.NewRow();

                for (int column = 1; column <= nColumn; column++)
                {
                    if (XlData[row, column]?.ToString() != "")
                    {
                        XlDataRow[column - 1] = XlData[row, column]?.ToString();
                    }
                }
                tmpXlData.Rows.Add(XlDataRow);
            }

            return tmpXlData;
        }

        public string OpenWkbk(string tmpwkbkFilePath)
        {
            bool originalDisplayAlerts = Globals.ThisAddIn.Application.DisplayAlerts;
            bool originalAskToUpdateLink = Globals.ThisAddIn.Application.AskToUpdateLinks;
            //Excel.XlCalculation originalAutoCalculation = Globals.ThisAddIn.Application.Calculation;

            Globals.ThisAddIn.Application.DisplayAlerts = false;
            Globals.ThisAddIn.Application.AskToUpdateLinks = false;
            //Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationManual;

            try
            {
                if (IfWkbkNotOpen(tmpwkbkFilePath))
                {
                    Excel.Workbook tmpWkbk = Globals.ThisAddIn.Application.Workbooks.Open(tmpwkbkFilePath, false, true);
                    tmpWkbk.Windows[1].Visible = false;
                    return tmpWkbk.Name.ToString();
                }

            }
            catch
            {
            }

            Globals.ThisAddIn.Application.DisplayAlerts = originalDisplayAlerts;
            Globals.ThisAddIn.Application.AskToUpdateLinks = originalAskToUpdateLink;
            return "";
            //Globals.ThisAddIn.Application.Calculation = originalAutoCalculation;
        }


        public string OpenWkbk(string tmpwkbkFilePath, Excel.Application myApp)
        {
            bool originalDisplayAlerts = myApp.DisplayAlerts;
            bool originalAskToUpdateLink = myApp.AskToUpdateLinks;
            //Excel.XlCalculation originalAutoCalculation = Globals.ThisAddIn.Application.Calculation;

            myApp.DisplayAlerts = false;
            myApp.AskToUpdateLinks = false;
            //Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationManual;

            try
            {
                if (IfWkbkNotOpen(tmpwkbkFilePath, myApp))
                {
                    Excel.Workbook tmpWkbk = myApp.Workbooks.Open(tmpwkbkFilePath, false, true);
                    tmpWkbk.Windows[1].Visible = false;
                    return tmpWkbk.Name.ToString();
                }

            }
            catch
            {
            }

            myApp.DisplayAlerts = originalDisplayAlerts;
            myApp.AskToUpdateLinks = originalAskToUpdateLink;
            return "";
            //Globals.ThisAddIn.Application.Calculation = originalAutoCalculation;
        }

        public DataTable SelectDistinct(DataTable InputTable, string[] columnName)
        {
            DataView myDataView = InputTable.DefaultView;

            DataTable outTable = myDataView.ToTable(/*distinct*/true, columnName);

            return outTable;
        }

        private bool IfWkbkNotOpen(string tmpwkbkFilePath)
        {
            bool NotOpen = true;

            foreach (Excel.Workbook Wkbk in Globals.ThisAddIn.Application.Workbooks)
            {
                if (tmpwkbkFilePath == Wkbk.FullName.ToString()) NotOpen = false;
            }

            return NotOpen;
        }

        private bool IfWkbkNotOpen(string tmpwkbkFilePath, Excel.Application myApp)
        {
            bool NotOpen = true;

            foreach (Excel.Workbook Wkbk in myApp.Workbooks)
            {
                if (tmpwkbkFilePath == Wkbk.FullName.ToString()) NotOpen = false;
            }

            return NotOpen;
        }
    }
}

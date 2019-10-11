using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;

namespace ReservingUploader
{
    class cls_SQL_ReservingFile
    {
        private Dictionary<string, double> FXRate;

        public DataTable GetWkbkParameterFromSQL()
        {
            string connectionStringSQL = @"Database=ADS;Server=CREREPSQL03;
                Integrated Security=True;connect timeout=60";

            //get SQL table as datatable
            DataTable myData = new DataTable();
            String SQLquery = "SELECT * FROM tmp_Reservingfiles_WkbkParameters";

            using (SqlConnection connectionSQL = new SqlConnection(connectionStringSQL))
            using (SqlCommand querySQL = new SqlCommand(SQLquery, connectionSQL))
            {
                try
                {
                    connectionSQL.Open();

                    // read DB to Datatable
                    SqlDataAdapter myDataAdapter = new SqlDataAdapter(querySQL);
                    myDataAdapter.Fill(myData);

                    connectionSQL.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
            return myData;
        }

        public DataTable ReservingFileData(DataTable parData, Excel.Application myApp)
        {
            DataTable DataTable = new DataTable();

            DataTable.Columns.Add("YOA", typeof(double));
            DataTable.Columns.Add("Financial Type", typeof(string));
            DataTable.Columns.Add("LOB", typeof(string));
            DataTable.Columns.Add("CB Class", typeof(string));
            DataTable.Columns.Add("Claim Type", typeof(string));
            DataTable.Columns.Add("Currency", typeof(string));
            DataTable.Columns.Add("RI Type", typeof(string));
            DataTable.Columns.Add("Value", typeof(double));
            DataTable.Columns.Add("TextValue", typeof(string));

            foreach (Excel.Workbook Wkbk in myApp.Workbooks)
            {
                if (Wkbk.Name.ToString().Split('_').Count() > 1)
                {
                    switch (Wkbk.Name.ToString().Split('_')[0].ToString() + "_"
                    + Wkbk.Name.ToString().Split('_')[1].ToString())
                    {
                        case "GAAP Gross Premium_Open Market":
                            for (int rowCount = 0; rowCount < parData.Rows.Count; rowCount++)
                            {
                                if (parData.Rows[rowCount][0].ToString() == "GAAP Gross Premium_Open Market")
                                {
                                    LoopWksht(DataTable,
                                    Wkbk,
                                    parData.Rows[rowCount][1].ToString(),
                                    Convert.ToInt32(parData.Rows[rowCount][2].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][3].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][4].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][5].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][6].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][7].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][8].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][9].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][10].ToString()),
                                    Convert.ToBoolean(parData.Rows[rowCount][11].ToString()));
                                }
                            }
                            break;
                        case "GAAP Gross Reserving_Open Market":
                            for (int rowCount = 0; rowCount < parData.Rows.Count; rowCount++)
                            {
                                if (parData.Rows[rowCount][0].ToString() == "GAAP Gross Reserving_Open Market")
                                {
                                    LoopWksht(DataTable,
                                    Wkbk,
                                    parData.Rows[rowCount][1].ToString(),
                                    Convert.ToInt32(parData.Rows[rowCount][2].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][3].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][4].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][5].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][6].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][7].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][8].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][9].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][10].ToString()),
                                    Convert.ToBoolean(parData.Rows[rowCount][11].ToString()));
                                }
                            }
                            break;
                        case "GAAP Gross Reserving_SPSPPV":
                            for (int rowCount = 0; rowCount < parData.Rows.Count; rowCount++)
                            {
                                if (parData.Rows[rowCount][0].ToString() == "GAAP Gross Reserving_SPSPPV")
                                {
                                    LoopWksht(DataTable,
                                    Wkbk,
                                    parData.Rows[rowCount][1].ToString(),
                                    Convert.ToInt32(parData.Rows[rowCount][2].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][3].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][4].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][5].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][6].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][7].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][8].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][9].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][10].ToString()),
                                    Convert.ToBoolean(parData.Rows[rowCount][11].ToString()));
                                }
                            }
                            break;
                        case "GAAP Reinsurance_Open Market and SPSPPV":
                            for (int rowCount = 0; rowCount < parData.Rows.Count; rowCount++)
                            {
                                if (parData.Rows[rowCount][0].ToString() == "GAAP Reinsurance_Open Market and SPSPPV")
                                {
                                    LoopWksht(DataTable,
                                    Wkbk,
                                    parData.Rows[rowCount][1].ToString(),
                                    Convert.ToInt32(parData.Rows[rowCount][2].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][3].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][4].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][5].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][6].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][7].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][8].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][9].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][10].ToString()),
                                    Convert.ToBoolean(parData.Rows[rowCount][11].ToString()));
                                }
                            }
                            break;
                        default:
                            break;
                    }
                }
            }
            return DataTable;
        }

        private void LoopWksht(
            DataTable upLoadTable,
            Excel.Workbook tgtWkbk,
            string TgtWkshtName,
            int L3Class,
            int CBClass,
            int YOA,
            int intScol,
            int intSrow,
            int FinancialType,
            int ClaimType,
            int RIType,
            int CurrencyType,
            bool Conv)
        {

            //Dictionary<string, double> FXRate = this.FXRate();

            Excel.Worksheet tgtWksht = tgtWkbk.Sheets[TgtWkshtName];

            int intEcol = tgtWksht.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
            int intErow = tgtWksht.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

            Excel.Range myRange = (Excel.Range)tgtWksht.Cells[1, 1].Resize[intErow, intEcol];

            object[,] myData = myRange.Value2;

            for (int tmpCol = intScol; tmpCol <= intEcol; tmpCol++)
            {
                for (int tmpRow = intSrow; tmpRow <= intErow; tmpRow++)
                {
                    try
                    {
                        if (myData[tmpRow, tmpCol] != null &&
                        Convert.ToString(myData[tmpRow, tmpCol]) != "x" &&
                        myData[tmpRow, YOA] != null &&
                        myData[FinancialType, tmpCol] != null &&
                        Convert.ToString(myData[tmpRow, tmpCol]) != "0" &&
                        Convert.ToString(myData[tmpRow, tmpCol]) != "-" &&
                        Convert.ToString(myData[tmpRow, YOA]) != "Total")
                        {
                            DataRow uploadRow = upLoadTable.NewRow();

                            uploadRow[0] = myData[tmpRow, YOA]; //YOA
                            uploadRow[1] = myData[FinancialType, tmpCol]; //Financial Type
                            if (myData[tmpRow, L3Class] != null)
                                uploadRow[2] = myData[tmpRow, L3Class]; //LOB
                            if (myData[tmpRow, CBClass] != null)
                                uploadRow[3] = myData[tmpRow, CBClass]; //CB Class;
                            if (myData[ClaimType, tmpCol] != null)
                                uploadRow[4] = myData[ClaimType, tmpCol]; //Claim Type
                            if (myData[CurrencyType, tmpCol] != null)
                                uploadRow[5] = myData[CurrencyType, tmpCol]; //Currency
                            if (myData[RIType, tmpCol] != null)
                                uploadRow[6] = myData[RIType, tmpCol]; //RI Type
                            if (Convert.ToString(myData[5, tmpCol]) == "TextValue")
                            {
                                uploadRow[8] = myData[tmpRow, tmpCol];
                            }
                            else
                            {
                                uploadRow[7] = myData[tmpRow, tmpCol];
                            }

                            if (uploadRow[6].ToString() != "G")
                                uploadRow[7] = Convert.ToDouble(uploadRow[7].ToString()) * -1;
                            if (Conv)
                                uploadRow[7] = Convert.ToDouble(uploadRow[7].ToString()) * FXRate[uploadRow[5].ToString()];

                            upLoadTable.Rows.Add(uploadRow);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString() + "Column: " + tmpCol + "Row: " + tmpRow);
                    }
                }
            }
            tgtWksht = null;
        }

        public void SetFXRate(Excel.Application myApp)
        {

            var tmpFXRate = new Dictionary<string, double>();

            foreach (Excel.Workbook Wkbk in myApp.Workbooks)
            {
                if (Wkbk.Name.ToString().Split('_').Count() > 1)
                {
                    if (Wkbk.Name.ToString().Split('_')[0].ToString() == "GAAP Gross Premium")
                    {
                        var CCYList = new List<string> { "USD", "GBP", "CAD", "EUR", "AUD", "JPY" };

                        foreach (string CCY in CCYList)
                        {
                            String itemName = "FXRates_Current_" + CCY;
                            tmpFXRate.Add(CCY,Convert.ToDouble(Wkbk.Names.Item(itemName).RefersToRange.Value2));
                        }
                    }
                }
            }
            
            FXRate=tmpFXRate;
        }
    }
}

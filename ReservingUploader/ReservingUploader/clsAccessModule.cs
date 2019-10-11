using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Access = Microsoft.Office.Interop.Access;
using DAO = Microsoft.Office.Interop.Access.Dao;
using System.IO;


namespace ReservingUploader
{
    class clsAccessModule
    {
        string QueryString;
        string ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;
                Data Source=U:\Reserving\Legacy Process\ADS_Uploader.accdb;";

        public DataTable AsAtData()
        {
            DataTable AsAtData = new DataTable();

            //change to line for now
            QueryString = @"SELECT DISTINCT [AsAt] FROM [tblVersionInfo]";

            using (OleDbConnection Connection = new OleDbConnection(ConnectionString))
            using (OleDbCommand Comm = new OleDbCommand(QueryString, Connection))
            {
                try
                {
                    Connection.Open();
                    OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(Comm);
                    myDataAdapter.Fill(AsAtData);
                    Connection.Close();
                }
                catch
                {
                }
            }

            return AsAtData;
        }

        public DataTable AccessData()
        {
            DataTable AccessData = new DataTable();

            //change to line for now
            QueryString = @"SELECT * FROM [tblCombinedDataToUpload]";


            using (OleDbConnection Connection = new OleDbConnection(ConnectionString))
            using (OleDbCommand Comm = new OleDbCommand(QueryString, Connection))
            {
                try
                {
                    Connection.Open();
                    OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(Comm);
                    myDataAdapter.Fill(AccessData);
                    Connection.Close();
                }
                catch
                {
                }
            }
            return AccessData;
        }

        public DataTable AccessData(String fileFullPath, String tableName)
        {
            DataTable AccessData = new DataTable();

            String QueryString = @"SELECT * FROM [" + tableName + "]";

            ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;
                Data Source="+ fileFullPath +";";

            //change to line for now
            //QueryString = @"SELECT * FROM [tblCombinedDataToUpload]";

            using (OleDbConnection Connection = new OleDbConnection(ConnectionString))
            using (OleDbCommand Comm = new OleDbCommand(QueryString, Connection))
            {
                try
                {
                    Connection.Open();
                    OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(Comm);
                    myDataAdapter.Fill(AccessData);
                    Connection.Close();
                }
                catch
                {
                }
            }
            return AccessData;
        }

        public DataTable resFileLocationData(string AsAt)
        {
            DataTable resFileLocationData = new DataTable();

            //change to line for now
            QueryString = @"SELECT [TargetWorkbook]";
            QueryString = QueryString + @" FROM [tblVersionInfo]";
            QueryString = QueryString + @" WHERE [AsAt]=" + @"""" + AsAt.ToString() + @"""";

            using (OleDbConnection Connection = new OleDbConnection(ConnectionString))
            using (OleDbCommand Comm = new OleDbCommand(QueryString, Connection))
            {
                try
                {
                    Connection.Open();
                    OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(Comm);
                    myDataAdapter.Fill(resFileLocationData);
                    Connection.Close();
                }
                catch
                {
                }
            }
            return resFileLocationData;
        }

        public Dictionary<string, double> FXRate()
        {
            DataTable FXRateData = new DataTable();

            //change to line for now
            QueryString = @"SELECT * FROM [tblFXRate]";

            using (OleDbConnection Connection = new OleDbConnection(ConnectionString))
            using (OleDbCommand Comm = new OleDbCommand(QueryString, Connection))
            {
                try
                {
                    Connection.Open();
                    OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(Comm);
                    myDataAdapter.Fill(FXRateData);
                    Connection.Close();
                }
                catch
                {
                }
            }

            var FXRate = new Dictionary<string, double>();

            for (int i = 0; i < FXRateData.Rows.Count; i++)
            {
                FXRate.Add(FXRateData.Rows[i][0].ToString(), Convert.ToDouble(FXRateData.Rows[i][1]));
            }

            return FXRate;
        }

        public DataTable GetParameterFromAccess()
        {
            string connString = @"Provider =Microsoft.ACE.OLEDB.12.0;
                Data Source=U:\Reserving\Legacy Process\ADS_Uploader.accdb;";

            DataTable result = new DataTable();

            using (OleDbConnection connection = new OleDbConnection(connString))
            using (OleDbCommand cmd = new OleDbCommand("SELECT * FROM tblParameter", connection))
            {
                try
                {
                    connection.Open();
                    OleDbDataAdapter adapter = new OleDbDataAdapter(cmd);
                    adapter.Fill(result);
                    connection.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
            return result;
        }

        public void UploadDataToAccess(DataTable sourceData, string DBPath, string TblName, bool ClearTbl)
        {
            Boolean CheckFl = false;
            DAO.DBEngine dbEngine = new DAO.DBEngine();

            try
            {
                DAO.Database db = dbEngine.OpenDatabase(DBPath);
                DAO.Recordset AccessRecordset = db.OpenRecordset(TblName);
                DAO.Field[] AccessFields = new DAO.Field[sourceData.Columns.Count];

                //Whether to clear table before pasting
                if (ClearTbl)
                    db.Execute("DELETE FROM " + TblName);

                for (Int32 rowCount = 0; rowCount < sourceData.Rows.Count; rowCount++)
                {
                    AccessRecordset.AddNew();
                    for (Int32 colCount = 0; colCount < sourceData.Columns.Count; colCount++)
                    {
                        if (!CheckFl)
                            AccessFields[colCount] = AccessRecordset.Fields[sourceData.Columns[colCount].ColumnName];
                        AccessFields[colCount].Value = sourceData.Rows[rowCount][colCount];
                    }
                    
                    try
                    {
                        AccessRecordset.Update();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString() + sourceData.Rows[rowCount][3].ToString());
                    }

                    CheckFl = true;
                }
                AccessRecordset.Close();
                db.Close();
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(dbEngine);
                dbEngine = null;
            }
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
            ///
            DataTable.Columns.Add("TextValue", typeof(string));
            ///

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
                                if (parData.Rows[rowCount][1].ToString() == "GAAP Gross Premium_Open Market")
                                {
                                    LoopWksht(DataTable,
                                    Wkbk,
                                    parData.Rows[rowCount][2].ToString(),
                                    Convert.ToInt32(parData.Rows[rowCount][3].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][4].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][5].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][6].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][7].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][8].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][9].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][10].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][11].ToString()),
                                    Convert.ToBoolean(parData.Rows[rowCount][12].ToString()));
                                }
                            }
                            break;
                        case "GAAP Gross Reserving_Open Market":
                            for (int rowCount = 0; rowCount < parData.Rows.Count; rowCount++)
                            {
                                if (parData.Rows[rowCount][1].ToString() == "GAAP Gross Reserving_Open Market")
                                {
                                    LoopWksht(DataTable,
                                    Wkbk,
                                    parData.Rows[rowCount][2].ToString(),
                                    Convert.ToInt32(parData.Rows[rowCount][3].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][4].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][5].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][6].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][7].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][8].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][9].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][10].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][11].ToString()),
                                    Convert.ToBoolean(parData.Rows[rowCount][12].ToString()));
                                }
                            }
                            break;
                        case "GAAP Gross Reserving_SPSPPV":
                            for (int rowCount = 0; rowCount < parData.Rows.Count; rowCount++)
                            {
                                if (parData.Rows[rowCount][1].ToString() == "GAAP Gross Reserving_SPSPPV")
                                {
                                    LoopWksht(DataTable,
                                    Wkbk,
                                    parData.Rows[rowCount][2].ToString(),
                                    Convert.ToInt32(parData.Rows[rowCount][3].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][4].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][5].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][6].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][7].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][8].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][9].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][10].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][11].ToString()),
                                    Convert.ToBoolean(parData.Rows[rowCount][12].ToString()));
                                }
                            }
                            break;
                        case "GAAP Reinsurance_Open Market and SPSPPV":
                            for (int rowCount = 0; rowCount < parData.Rows.Count; rowCount++)
                            {
                                if (parData.Rows[rowCount][1].ToString() == "GAAP Reinsurance_Open Market and SPSPPV")
                                {
                                    LoopWksht(DataTable,
                                    Wkbk,
                                    parData.Rows[rowCount][2].ToString(),
                                    Convert.ToInt32(parData.Rows[rowCount][3].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][4].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][5].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][6].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][7].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][8].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][9].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][10].ToString()),
                                    Convert.ToInt32(parData.Rows[rowCount][11].ToString()),
                                    Convert.ToBoolean(parData.Rows[rowCount][12].ToString()));
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

        public string CloseResWkbk()
        {
            foreach (Excel.Workbook Wkbk in Globals.ThisAddIn.Application.Workbooks)
            {
                if (Wkbk.Name.ToString().Split('_').Count() > 1 &&
                    Wkbk.Name.ToString().Substring(0, 4) == "GAAP")
                {
                    Wkbk.Close(false);
                    return Wkbk.Name;
                }
            }

            return "";
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
            Dictionary<string, double> FXRate = this.FXRate();

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


        public void runAccessQuery()
        {
            Access.Application myAccess = new Access.Application();

            myAccess.Visible = false;

            myAccess.OpenCurrentDatabase(@"U:\Reserving\Legacy Process\ADS_Uploader.accdb");

            myAccess.DoCmd.RunSQL(@"DELETE * FROM tblCombinedDataToUpload");

            myAccess.DoCmd.OpenQuery(@"qryOutputData");

            //myAccess.DoCmd.RunMacro("Test");

            myAccess.Quit(Access.AcQuitOption.acQuitSaveNone);
        }
    }
}

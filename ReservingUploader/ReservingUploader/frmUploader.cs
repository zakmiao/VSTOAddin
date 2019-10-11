using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.Data.OleDb;
using Access = Microsoft.Office.Interop.Access;


namespace ReservingUploader
{
    public partial class frmUploader : Form
    {
        public frmUploader()
        {
            InitializeComponent();
        }

        //load common objects
        DataTable myData = new DataTable();
        clsAccessModule AccessModule = new clsAccessModule();
        cls_SQL_ReservingFile SQLResModule = new cls_SQL_ReservingFile();
        clsSQLModule SQLModule = new clsSQLModule();

        private String guessFileType(string FileName)
        {
            //use regular expression to judge reserving file or majorevent file
            String resPattern = "^GAAP.*$";
            String majPattern = "^Major Events.*$";
            String ClmBPattern = "^.*Claimsband.*$";
            String ClaimsListPattern = "^Claims List.*$";
            String result = "";

            if (System.Text.RegularExpressions.Regex.IsMatch(FileName.ToString(), resPattern))
                return "Reserving File";

            if (System.Text.RegularExpressions.Regex.IsMatch(FileName.ToString(), majPattern))
                return "MajorEvent File";

            if (System.Text.RegularExpressions.Regex.IsMatch(FileName.ToString(), ClmBPattern))
                return "ClaimsBand File";

            if (System.Text.RegularExpressions.Regex.IsMatch(FileName.ToString(), ClaimsListPattern))
                return "ClaimsList File";

            return result;
        }

        private String guessAsAt(string FullName)
        {
            //use regular expression to judge file asat from file location
            String yrPattern = "^20[0-1][0-9]$";

            String moPattern = "^(m|M|q|Q)[1]?[0-9]";

            String[] words = FullName.Split('\\');

            String result = "";

            foreach (var word in words)
            {
                if (System.Text.RegularExpressions.Regex.IsMatch(word.ToString(), yrPattern)
                    || System.Text.RegularExpressions.Regex.IsMatch(word.ToString(), moPattern))
                {
                    result = result + word.ToString().ToUpper();
                }
            }
            return result;
        }

        private Int32 UniqueAsAt(DataTable myTable)
        {
            DataView myView = new DataView(myTable);
            DataTable distinctValue = myView.ToTable(true, "AsAt");
            return distinctValue.Rows.Count;
        }

        private void UpdataFXRateQry(Excel.Workbook myWkbk, string CCY, OleDbConnection Connection)
        {
            string itemName = "FXRates_Current_" + CCY;

            String QueryString = @"UPDATE [tblFXRate] SET [FXRate]="
                            + myWkbk.Names.Item(itemName).RefersToRange.Value2
                            + @" WHERE [CCY]='" + CCY + "'";

            OleDbCommand Comm = new OleDbCommand(QueryString, Connection);

            Comm.ExecuteNonQuery();
        }

        private void ExtractClaimsBandData(Excel.Worksheet tmpSheet, DataTable CBData, Boolean Claim, Boolean RI)
        {
            Excel.Range myRange = tmpSheet.UsedRange;

            object[,] XlData = myRange.Value2;

            var nColumn = myRange.Columns.Count;
            var nRow = myRange.Rows.Count;

            for (int column = 4; column <= 10; column++)
            {
                if (new[] { "GBP", "USD", "CAD", "EUR", "AUD", "JPY" }.Contains(XlData[1, column]?.ToString()))
                {
                    for (int row = 2; row <= nRow; row++)
                    {
                        DataRow XlDataRow = CBData.NewRow();

                        try
                        {
                            if (XlData[row, column] != null)
                            {
                                if (Convert.ToDouble(XlData[row, column]?.ToString()) != 0)
                                {
                                    XlDataRow["AsAt"] = myData.Rows[0][1].ToString();
                                    XlDataRow["UWY"] = Convert.ToInt32(XlData[row, 1]?.ToString());
                                    XlDataRow["DataType"] = XlData[row, 2]?.ToString();

                                    if (Claim)
                                    {
                                        XlDataRow["ClaimType"] = XlData[row, 3]?.ToString();
                                        XlDataRow["SBFClass"] = XlData[row, 5]?.ToString();
                                    }
                                    else
                                    {
                                        XlDataRow["ClaimType"] = "PREM";
                                        XlDataRow["SBFClass"] = XlData[row, 4]?.ToString();
                                    }

                                    if (RI) { XlDataRow["RIType"] = "C1"; }
                                    else { XlDataRow["RIType"] = "G"; }
                                    XlDataRow["Currency"] = XlData[1, column]?.ToString();
                                    XlDataRow["Value"] = Convert.ToDouble(XlData[row, column]?.ToString());

                                    CBData.Rows.Add(XlDataRow);
                                }
                            }
                        }
                        catch
                        {
                            MessageBox.Show(column.ToString() + " column & row" + row.ToString());
                        }
                    }

                }
            }
        }

        public string OpenWkbk(string tmpwkbkFilePath, Excel.Application myApp)
        {
            bool originalDisplayAlerts = myApp.DisplayAlerts;
            bool originalAskToUpdateLink = myApp.AskToUpdateLinks;

            myApp.DisplayAlerts = false;
            myApp.AskToUpdateLinks = false;

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

        private void frmUploader_Load(object sender, EventArgs e)
        {
            //set up the datatable to show in the datagridview
            myData.Columns.Add("FilePath", typeof(System.String));
            myData.Columns["FilePath"].ReadOnly = true;
            myData.Columns.Add("AsAt", typeof(System.String));
            myData.Columns.Add("FileType", typeof(System.String));

            //initial message on statusbar
            this.toolStripStatusLabel1.Text = "Drag and Drop the files to upload.";
            this.TopMost = true;
        }

        private void FileDropDataGridView_DragDrop(object sender, DragEventArgs e)
        {
            //action: drag and drop into datagridview
            try
            {
                //get file fullname
                String[] files = (String[])e.Data.GetData(DataFormats.FileDrop);

                foreach (String file in files)
                {
                    //get file attributes
                    System.IO.FileInfo fi = new System.IO.FileInfo(file);
                    Shell32.Shell fileshell = new Shell32.Shell();
                    Shell32.Folder fileshellFolder = fileshell.NameSpace(fi.Directory.ToString() + @"\");
                    Shell32.FolderItem fileshellItem = fileshellFolder.ParseName(fi.Name);

                    //show error message when not res or me file
                    if (guessFileType(fi.Name).ToString() == "")
                    {
                        MessageBox.Show(this, "Error: file can not be uploaded.");
                        return;
                    }

                    //add file information into datagridview
                    DataRow myRow = myData.NewRow();
                    myRow["FilePath"] = file;
                    myRow["AsAt"] = guessAsAt(files[0]).ToString();
                    myRow["fileType"] = guessFileType(fi.Name).ToString();

                    myData.Rows.Add(myRow);
                    this.toolStripStatusLabel1.Text = "Double click left side to delete row, edit AsAt and Filetype if incorrect.";
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            this.FileDropDataGridView.DataSource = myData;
            this.FileDropDataGridView.AutoResizeColumns();
        }

        private void FileDropDataGridView_DragEnter(object sender, DragEventArgs e)
        {
            //standard: dragenter event
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.All;
            else
                e.Effect = DragDropEffects.None;
        }

        private void FileDropDataGridView_RowHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            //event: double click to remove row
            this.FileDropDataGridView.Rows.Remove(this.FileDropDataGridView.Rows[e.RowIndex]);
        }

        private void UploadButton_Click(object sender, EventArgs e)
        {
            
            //main: open files, fill datatable, upload to SQL tmp database
            DataTable myData = (DataTable)this.FileDropDataGridView.DataSource;
            Boolean ResExist = false;
            Boolean MEExist = false;
            Boolean ClaimListExist = false;

            if (UniqueAsAt(myData) > 1)
            {
                MessageBox.Show(this, "Error: AsAt should be the same for each file.");
                return;
            }

            //load the version control userform with all versions for that asat choosen
            frmUploader_VersionControl testDialog = new frmUploader_VersionControl();
            testDialog.StartPosition = FormStartPosition.CenterScreen;
            testDialog.GetVersion(myData.Rows[0][1].ToString());
            testDialog.TopMost = true;

            if (testDialog.ShowDialog(this) == DialogResult.OK)
            {
                //process start
                this.toolStripStatusLabel1.Text = "Uploading";

                //open new excel application to accomendate all the spreadsheets
                //*********************************************************************************************************
                Excel.Application myApp = new Excel.Application();
                myApp.Visible = false;

                //Excel.Application myApp = Globals.ThisAddIn.Application;

                //*********************************************************************************************************

                //set up the major event file
                Excel.Workbook MEWkbk = null;

                //set up the Claimsband file
                Excel.Workbook CBWkbk = null;

                //set up a reserving file
                Excel.Workbook ResWkbk = null;

                //set up a ClaimsList file
                Excel.Workbook ClaimsListWkbk = null;

                for (int i = 0; i < myData.Rows.Count; i++)
                {
                    string openWkbkStatus = OpenWkbk(myData.Rows[i][0].ToString(), myApp);

                    if (myData.Rows[i][2].ToString() == "MajorEvent File")
                    {
                        MEWkbk = myApp.Workbooks[myData.Rows[i][0].ToString().Split('\\').Last()];
                    }

                    if (myData.Rows[i][2].ToString() == "ClaimsBand File")
                    {
                        CBWkbk = myApp.Workbooks[myData.Rows[i][0].ToString().Split('\\').Last()];
                    }

                    if (myData.Rows[i][2].ToString() == "Reserving File")
                    {
                        ResWkbk = myApp.Workbooks[myData.Rows[i][0].ToString().Split('\\').Last()];
                    }

                    if (myData.Rows[i][2].ToString() == "ClaimsList File")
                    {
                        ClaimsListWkbk = myApp.Workbooks[myData.Rows[i][0].ToString().Split('\\').Last()];
                    }

                    //progressbar
                    this.toolStripProgressBar1.Value = 40 * (i + 1) / (myData.Rows.Count);
                }

                //progressbar
                this.toolStripProgressBar1.Value = 40;

                //recalculate before populate datatable
                myApp.Calculate();

                //progressbar
                this.toolStripProgressBar1.Value = 50;

                //set up datatable to collect data from excel files
                DataTable ResData = new DataTable();
                DataTable MEData = new DataTable();
                DataTable CBData = new DataTable();
                DataTable ClaimsListData = new DataTable();

                // Reserving files
                if (ResWkbk != null)
                {
                    ResExist = true;

                    //get the parameters for tab name, starting rows, starting columns...
                    DataTable parData = SQLResModule.GetWkbkParameterFromSQL();

                    //set up FX rates from reservering files
                    SQLResModule.SetFXRate(myApp);

                    //populate reserving datatable
                    ResData = SQLResModule.ReservingFileData(parData, myApp);

                    //Add AsAt Column
                    DataColumn AsAtCol = new DataColumn("AsAt");
                    AsAtCol.DefaultValue = myData.Rows[0][1].ToString();
                    ResData.Columns.Add(AsAtCol);
                    AsAtCol.SetOrdinal(0);

                    //Add Version Column
                    DataColumn VersionCol = new DataColumn("Version");
                    VersionCol.DefaultValue = testDialog.Version.ToString();
                    ResData.Columns.Add(VersionCol);
                    VersionCol.SetOrdinal(1);
                    
                    //upload raw data to SQL
                    uploadToSQL(ResData, "tmp_Reservingfiles_InputData");
                }
                
                //prograss bar
                this.toolStripProgressBar1.Value = 60;

                //Major Event Data
                DataTable MEuploadData = new DataTable();

                if (MEWkbk != null)
                {
                    MEExist = true;

                    try
                    {
                        Excel.Worksheet mySheet = MEWkbk.Worksheets["ME FlatFile"];
                        Excel.Range myRange = mySheet.UsedRange;

                        object[,] XlData = myRange.Value2;
                        
                        MEData.Columns.Add("Event", System.Type.GetType("System.String"));
                        MEData.Columns.Add("Cedant", System.Type.GetType("System.String"));
                        MEData.Columns.Add("UWRef", System.Type.GetType("System.String"));
                        MEData.Columns.Add("UWY", System.Type.GetType("System.Int32"));
                        MEData.Columns.Add("DataType", System.Type.GetType("System.String"));
                        MEData.Columns.Add("SBFClass", System.Type.GetType("System.String"));
                        MEData.Columns.Add("ClaimType", System.Type.GetType("System.String"));
                        MEData.Columns.Add("RIType", System.Type.GetType("System.String"));
                        MEData.Columns.Add("Currency", System.Type.GetType("System.String"));
                        MEData.Columns.Add("Value", System.Type.GetType("System.Double"));
                        MEData.Columns.Add("TextValue", System.Type.GetType("System.String"));

                        var nColumn = myRange.Columns.Count;
                        var nRow = myRange.Rows.Count;

                        for (int row = 2; row <= nRow; row++)
                        {
                            DataRow XlDataRow = MEData.NewRow();
                            
                            XlDataRow["Event"] = XlData[row, 2]?.ToString();
                            XlDataRow["Cedant"] = XlData[row, 3]?.ToString();
                            XlDataRow["UWRef"] = XlData[row, 4]?.ToString();
                            XlDataRow["UWY"] = Convert.ToInt32(XlData[row, 5]?.ToString());
                            XlDataRow["DataType"] = XlData[row, 6]?.ToString();
                            XlDataRow["SBFClass"] = XlData[row, 7]?.ToString();
                            XlDataRow["ClaimType"] = XlData[row, 8]?.ToString();
                            XlDataRow["RIType"] = XlData[row, 9]?.ToString();
                            XlDataRow["Currency"] = XlData[row, 10]?.ToString();
                            XlDataRow["Value"] = Convert.ToDouble(XlData[row, 11]?.ToString());
                            XlDataRow["TextValue"] = XlData[row, 12]?.ToString();

                            MEData.Rows.Add(XlDataRow);
                        }

                        //Add AsAt Column
                        DataColumn AsAtCol = new DataColumn("AsAt");
                        AsAtCol.DefaultValue = myData.Rows[0][1].ToString();
                        MEData.Columns.Add(AsAtCol);
                        AsAtCol.SetOrdinal(0);

                        //Add Version Column
                        DataColumn VersionCol = new DataColumn("Version");
                        VersionCol.DefaultValue = testDialog.Version.ToString();
                        MEData.Columns.Add(VersionCol);
                        VersionCol.SetOrdinal(1);
                        
                        SQLModule.UploadMEToSQL(MEData, "tmp_ME_InputData");

                        SQLModule.MESQLStoredProcedure("sp_tmp_ME_PaidOS", testDialog.Version.ToString(), myData.Rows[0][1].ToString());
                        
                    }
                    catch
                    {
                    }
                    
                }

                //prograss bar
                this.toolStripProgressBar1.Value = 70;

                //Claim Band part

                /*
                if (CBWkbk != null)
                {
                    CBData.Columns.Add("AsAt", System.Type.GetType("System.String"));
                    CBData.Columns.Add("Event", System.Type.GetType("System.String"));
                    CBData.Columns.Add("Cedant", System.Type.GetType("System.String"));
                    CBData.Columns.Add("UWRef", System.Type.GetType("System.String"));
                    CBData.Columns.Add("UWY", System.Type.GetType("System.Int32"));
                    CBData.Columns.Add("DataType", System.Type.GetType("System.String"));
                    CBData.Columns.Add("SBFClass", System.Type.GetType("System.String"));
                    CBData.Columns.Add("ClaimType", System.Type.GetType("System.String"));
                    CBData.Columns.Add("RIType", System.Type.GetType("System.String"));
                    CBData.Columns.Add("Currency", System.Type.GetType("System.String"));
                    CBData.Columns.Add("Value", System.Type.GetType("System.Double"));
                    CBData.Columns.Add("TextValue", System.Type.GetType("System.String"));

                    try
                    {
                        foreach (Excel.Worksheet tmpSheet in CBWkbk.Worksheets)
                        {
                            switch (tmpSheet.Name.ToString())
                            {
                                case "Ultimates":
                                case "Gross Premiums":
                                    ExtractClaimsBandData(tmpSheet, CBData, false, false);
                                    break;

                                case "Gross Claims":
                                    ExtractClaimsBandData(tmpSheet, CBData, true, false);
                                    break;

                                case "RI Premiums":
                                    ExtractClaimsBandData(tmpSheet, CBData, false, true);
                                    break;

                                case "RI Claims":
                                    ExtractClaimsBandData(tmpSheet, CBData, true, true);
                                    break;

                                default:
                                    break;
                            }
                        }

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }

                    AccessModule.UploadDataToAccess(CBData, @"U:\Reserving\Claims Band\Claims Band.accdb", @"ClaimsBand_FlatFile", true);
                }
                */


                //ClaimsList part
                if (ClaimsListWkbk != null)
                {
                    ClaimListExist = true;

                    ClaimsListData.Columns.Add("SBFClass", System.Type.GetType("System.String"));
                    ClaimsListData.Columns.Add("UWRef", System.Type.GetType("System.String"));
                    ClaimsListData.Columns.Add("Bureau Claim Assured", System.Type.GetType("System.String"));
                    ClaimsListData.Columns.Add("Risk Code", System.Type.GetType("System.String"));
                    ClaimsListData.Columns.Add("UWY", System.Type.GetType("System.Int32"));
                    ClaimsListData.Columns.Add("Claim Ref", System.Type.GetType("System.String"));
                    ClaimsListData.Columns.Add("Claim Status", System.Type.GetType("System.String"));
                    ClaimsListData.Columns.Add("EventCode", System.Type.GetType("System.Int32"));
                    ClaimsListData.Columns.Add("Date of Loss", System.Type.GetType("System.String"));
                    ClaimsListData.Columns.Add("Date Last Movement", System.Type.GetType("System.String"));
                    ClaimsListData.Columns.Add("Movement - Date Entered", System.Type.GetType("System.String"));
                    ClaimsListData.Columns.Add("Reference", System.Type.GetType("System.String"));
                    ClaimsListData.Columns.Add("Loss Title", System.Type.GetType("System.String"));
                    ClaimsListData.Columns.Add("Current Narrative", System.Type.GetType("System.String"));
                    ClaimsListData.Columns.Add("Currency", System.Type.GetType("System.String"));
                    ClaimsListData.Columns.Add("Trust Fund Code", System.Type.GetType("System.String"));
                    ClaimsListData.Columns.Add("Paid Claims", System.Type.GetType("System.Double"));
                    ClaimsListData.Columns.Add("OS Claims", System.Type.GetType("System.Double"));

                    Excel.Range myClaimListRange = ClaimsListWkbk.Sheets["Eclipse Data"].UsedRange;

                    object[,] ClaimsListSheetData = myClaimListRange.Value2;

                    var nClaimListColumn = myClaimListRange.Columns.Count;
                    var nClaimListRow = myClaimListRange.Rows.Count;


                    for (int row = 2; row <= nClaimListRow-6; row++)
                    {
                        if (ClaimsListSheetData[row, 1]?.ToString() != ""
                            && ClaimsListSheetData[row, 1]?.ToString() != "Totals")
                        {
                            DataRow newRow = ClaimsListData.NewRow();

                            newRow["SBFClass"] = ClaimsListSheetData[row, 1]?.ToString();
                            newRow["UWRef"] = ClaimsListSheetData[row, 3]?.ToString();
                            newRow["Bureau Claim Assured"] = ClaimsListSheetData[row, 4]?.ToString();
                            newRow["Risk Code"] = ClaimsListSheetData[row, 5]?.ToString();
                            newRow["UWY"] = Convert.ToInt32(ClaimsListSheetData[row, 6]?.ToString());
                            newRow["Claim Ref"] = ClaimsListSheetData[row, 7]?.ToString();
                            newRow["Claim Status"] = ClaimsListSheetData[row, 8]?.ToString();
                            newRow["EventCode"] = Convert.ToInt32(ClaimsListSheetData[row, 9]?.ToString());
                            if (ClaimsListSheetData[row, 10]?.ToString() != "")
                            {
                                if(Convert.ToDouble(ClaimsListSheetData[row, 10]?.ToString()) > 0)
                                {
                                    DateTime DateLossTo = DateTime.FromOADate(Convert.ToDouble(ClaimsListSheetData[row, 10]?.ToString()));
                                    newRow["Date of Loss"] = DateLossTo;
                                }
                            }
                            if (ClaimsListSheetData[row, 12]?.ToString() != "")
                            {
                                DateTime DateLossTo = DateTime.FromOADate(Convert.ToDouble(ClaimsListSheetData[row, 12]?.ToString()));
                                newRow["Date Last Movement"] = DateLossTo;
                            }
                            if (ClaimsListSheetData[row, 13]?.ToString() != "")
                            {
                                DateTime DateLossTo = DateTime.FromOADate(Convert.ToDouble(ClaimsListSheetData[row, 13]?.ToString()));
                                newRow["Movement - Date Entered"] = DateLossTo;
                            }
                            newRow["Reference"] = ClaimsListSheetData[row, 14]?.ToString();
                            newRow["Loss Title"] = ClaimsListSheetData[row, 15]?.ToString();
                            newRow["Current Narrative"] = ClaimsListSheetData[row, 16]?.ToString();
                            newRow["Currency"] = ClaimsListSheetData[row, 17]?.ToString();
                            newRow["Trust Fund Code"] = ClaimsListSheetData[row, 18]?.ToString();
                            newRow["Paid Claims"] = Convert.ToDouble(ClaimsListSheetData[row, 19]?.ToString())
                                + Convert.ToDouble(ClaimsListSheetData[row, 20]?.ToString());
                            newRow["OS Claims"] = Convert.ToDouble(ClaimsListSheetData[row, 21]?.ToString())
                                + Convert.ToDouble(ClaimsListSheetData[row, 22]?.ToString());

                            ClaimsListData.Rows.Add(newRow);
                        }
                    }
                                        
                    ClaimListToSQL(ClaimsListData);

                    //call stored procedure in SQL
                    using (SqlConnection connectionSQL = new SqlConnection(@"Database=ADS;Server=CREREPSQL03;Integrated Security=True;connect timeout=30"))
                    {
                        connectionSQL.Open();

                        SqlCommand storedProcedureComm = new SqlCommand("sp_ClaimList_0_RunAll", connectionSQL);
                        storedProcedureComm.CommandType = CommandType.StoredProcedure;
                        try
                        {
                            storedProcedureComm.ExecuteNonQuery();
                        }
                        catch
                        {
                            MessageBox.Show("A problem with sp_ClaimList_0_RunAll.");
                        }

                        connectionSQL.Close();
                    }
                }

                //prograss bar
                this.toolStripProgressBar1.Value = 80;
                

                //upload various table into imp_tmp_UploadData
                using (SqlConnection connectionSQL = new SqlConnection(@"Database=ADS;Server=CREREPSQL03;Integrated Security=True;connect timeout=30"))
                {
                    connectionSQL.Open();

                    SqlCommand storedProcedureComm = new SqlCommand("sp_tmp_upd_3_uploadData", connectionSQL);
                    storedProcedureComm.CommandType = CommandType.StoredProcedure;
                    try
                    {
                        storedProcedureComm.Parameters.Add("@ResExist", SqlDbType.Bit).Value = ResExist;
                        storedProcedureComm.Parameters.Add("@MEExist", SqlDbType.Bit).Value = MEExist;
                        storedProcedureComm.Parameters.Add("@CLExist", SqlDbType.Bit).Value = ClaimListExist;
                        storedProcedureComm.ExecuteNonQuery();
                        
                    }
                    catch(Exception ex)
                    {
                        MessageBox.Show(this, ex.ToString() + @"A problem with sp_tmp_upd_3_uploadData.");
                    }

                    connectionSQL.Close();
                }

                this.toolStripProgressBar1.Value = 90;


                //close all excel spreadsheets
                foreach (Excel.Workbook Wkbk in myApp.Workbooks)
                {

                    Wkbk.Close(false);
                }


                //quite the excel application without saving

                //**********************************************************************************************
                myApp.Quit();
                //**********************************************************************************************


                //prograss bar
                this.toolStripProgressBar1.Value = 100;

                this.Activate();
                this.toolStripStatusLabel1.Text = "Upload success";
            }

            //clear up
            testDialog.Dispose();
        }

        private void ClaimListToSQL(DataTable ClaimListToBeProcessed)
        {
            // Upload amended excel sheet to SQL table
            //DataTable myUpTable = ObjSQLSchemaData("Columns", "ADS", this.tableListComboBox.SelectedItem.ToString());
            
            //connection string
            string connectionStringSQL = @"Database=ADS;Server=CREREPSQL03;
                Integrated Security=True;connect timeout=30";

            //query string to drop TestTab if exist
            string queryStringSQL = "IF OBJECT_ID('dbo.ClaimList_InputData') IS NOT NULL DROP TABLE dbo.ClaimList_InputData";
            queryStringSQL = queryStringSQL + @" CREATE TABLE ClaimList_InputData([SBFClass] [nvarchar](50) NULL," +
                @" [UWRef][nvarchar](25) NULL, [Bureau Claim Assured] [nvarchar] (100) NULL, [Risk Code] [nchar] (10) NULL," +
                @" [UWY] [int] NULL, [Claim Ref] [nvarchar] (50) NULL, [Claim Status] [nchar] (10) NULL, [EventCode] [int] NULL, " +
                @"[Date of Loss] [date] NULL, [Date Last Movement] [date] NULL, [Movement - Date Entered] [datetime] NULL, [Reference] [nvarchar] (50) NULL, " +
                @"[Loss Title] [nvarchar] (100) NULL, [Current Narrative] [nvarchar] (250) NULL, [Currency] [nchar] (10) NULL, " +
                @"[Trust Fund Code] [nchar] (10) NULL, [Paid Claims] [float] NULL,[OS Claims] [float] NULL)";
            
            //upload
            using (SqlConnection connectionSQL = new SqlConnection(connectionStringSQL))
            using (SqlCommand querySQL = new SqlCommand(queryStringSQL, connectionSQL))
            using (SqlBulkCopy bulkCopySQL = new SqlBulkCopy(connectionSQL))
            {
                connectionSQL.Open();

                //delete TestTab
                querySQL.ExecuteNonQuery();

                //copy excel sheet to SQL table
                bulkCopySQL.DestinationTableName = "ClaimList_InputData";
                bulkCopySQL.WriteToServer(ClaimListToBeProcessed);

                connectionSQL.Close();
            }
        }

        private void uploadToSQL(DataTable tableToUpload, String destTable)
        {
            // Upload to SQL table
            
            //connection string
            string connectionStringSQL = @"Database=ADS;Server=CREREPSQL03;
                Integrated Security=True;connect timeout=60";

            String queryStringSQL = @"DELETE FROM " + destTable;

            //upload
            using (SqlConnection connectionSQL = new SqlConnection(connectionStringSQL))
            using (SqlCommand querySQL = new SqlCommand(queryStringSQL, connectionSQL))
            using (SqlBulkCopy bulkCopySQL = new SqlBulkCopy(connectionSQL))
            {
                connectionSQL.Open();

                //delete table
                querySQL.ExecuteNonQuery();

                //upload data table to SQL
                bulkCopySQL.DestinationTableName = destTable;
                bulkCopySQL.WriteToServer(tableToUpload);

                connectionSQL.Close();
            }
        }
    }
}

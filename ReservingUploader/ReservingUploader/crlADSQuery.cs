using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;

namespace ReservingUploader
{
    public partial class crlADSQuery : UserControl
    {
        //Create SQL connection string
        string connectionStringSQL = @"Database=ADS;Server=CREREPSQL03;
                Integrated Security=True;connect timeout=30";

        //Create SQL quary string
        string queryStringSQL;

        public DataTable ObjSQLSchemaData(string Schema)
        {
            //get tables names
            DataTable schemaTable = new DataTable();

            using (SqlConnection connectionSQL = new SqlConnection(connectionStringSQL))
            {
                try
                {
                    connectionSQL.Open();
                    schemaTable = connectionSQL.GetSchema(Schema);
                    connectionSQL.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            return schemaTable;
        }

        public DataTable ObjSQLData(string SQLquery)
        {
            //get SQL table as datatable
            DataTable myData = new DataTable();

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
                    MessageBox.Show(ex.Message);
                }
            }
            return myData;
        }

        private String ViewSQLString()
        {
            string SQLString = "";

            if (criteriaComboBox1.SelectedItem != null
                && criteriatextBox1.Text != null)
            {
                SQLString = SQLString + '[' + criteriaComboBox1.SelectedItem.ToString() + ']' + ' ' + this.criteriatextBox1.Text;
            }

            if (criteriaComboBox2.SelectedItem != null
                && this.criteriatextBox2.Text != null)
            {
                if (criteriaComboBox2.SelectedItem.ToString() != "")
                    SQLString = SQLString + " AND [" + criteriaComboBox2.SelectedItem.ToString() + ']' + ' ' + this.criteriatextBox2.Text;
            }

            if (criteriaComboBox3.SelectedItem != null
                && this.criteriatextBox3.Text != null)
            {
                if (criteriaComboBox3.SelectedItem.ToString() != "")
                    SQLString = SQLString + " AND [" + criteriaComboBox3.SelectedItem.ToString() + ']' + ' ' + this.criteriatextBox3.Text;
            }

            if (criteriaComboBox4.SelectedItem != null
                && this.criteriatextBox4.Text != null)
            {
                if (criteriaComboBox4.SelectedItem.ToString() != "")
                    SQLString = SQLString + " AND [" + criteriaComboBox4.SelectedItem.ToString() + ']' + ' ' + this.criteriatextBox4.Text;
            }

            if (criteriaComboBox5.SelectedItem != null
                && this.criteriatextBox5.Text != null)
            {
                if (criteriaComboBox5.SelectedItem.ToString() != "")
                    SQLString = SQLString + " AND [" + criteriaComboBox5.SelectedItem.ToString() + ']' + ' ' + this.criteriatextBox5.Text;
            }

            return SQLString;
        }

        public DataTable ObjSQLSchemaData(string Schema, string DBName, string tabName)
        {
            //get columns names
            DataTable schemaTable = new DataTable();

            using (SqlConnection connectionSQL = new SqlConnection(connectionStringSQL))
            {
                try
                {
                    connectionSQL.Open();
                    schemaTable = connectionSQL.GetSchema(Schema, new[] { DBName, null, tabName });
                    connectionSQL.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            return schemaTable;
        }

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
            Excel.Worksheet mySheet = Globals.ThisAddIn.Application.ActiveSheet;
            mySheet.Cells.Clear();
            try
            {
                mySheet.Range["A1"].Resize[numRows + 1, numCols].Value = myArray;
                mySheet.Range["A1"].Resize[numRows + 1, numCols].Columns.AutoFit();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        public DataTable ExcelToDatatable()
        {
            DataTable tmpXlData = new DataTable();
            Excel.Worksheet mySheet = Globals.ThisAddIn.Application.ActiveSheet;
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

        public crlADSQuery()
        {
            InitializeComponent();
        }

        private void ViewRefresh_Click(object sender, EventArgs e)
        {
            DataTable myTmpSchemaData = ObjSQLSchemaData("Tables");
            viewListComboBox.Items.Clear();

            var numRows = myTmpSchemaData.Rows.Count;
            for (var row = 0; row < numRows; row++)
            {
                var tmpTableType = myTmpSchemaData.Rows[row][3].ToString();
                if (tmpTableType == "VIEW")
                    viewListComboBox.Items.Add(myTmpSchemaData.Rows[row][2]);
            }

            //clear criteria
            criteriaComboBox1.Items.Clear();
            this.criteriaComboBox1.SelectedItem = null;
            this.criteriaComboBox1.Text = "";
            this.criteriatextBox1.Text = "";

            criteriaComboBox2.Items.Clear();
            this.criteriaComboBox2.SelectedItem = null;
            this.criteriaComboBox2.Text = "";
            this.criteriatextBox2.Text = "";

            criteriaComboBox3.Items.Clear();
            this.criteriaComboBox3.SelectedItem = null;
            this.criteriaComboBox3.Text = "";
            this.criteriatextBox3.Text = "";

            criteriaComboBox4.Items.Clear();
            this.criteriaComboBox4.SelectedItem = null;
            this.criteriaComboBox4.Text = "";
            this.criteriatextBox4.Text = "";

            criteriaComboBox5.Items.Clear();
            this.criteriaComboBox5.SelectedItem = null;
            this.criteriaComboBox5.Text = "";
            this.criteriatextBox5.Text = "";
        }

        private void tableRefresh_Click(object sender, EventArgs e)
        {
            DataTable myTmpSchemaData = ObjSQLSchemaData("Tables");
            tableListComboBox.Items.Clear();

            var numRows = myTmpSchemaData.Rows.Count;
            for (var row = 0; row < numRows; row++)
            {
                var tmpTableType = myTmpSchemaData.Rows[row][3].ToString();
                if (tmpTableType == "BASE TABLE"
                    && myTmpSchemaData.Rows[row][2].ToString().Substring(0, 3) != "dat")
                    tableListComboBox.Items.Add(myTmpSchemaData.Rows[row][2]);
            }
        }

        private void viewDownload_Click(object sender, EventArgs e)
        {
            //set all objects to null
            queryStringSQL = "";
            DataTable myData = new DataTable();

            //set SQL query
            if (viewListComboBox.SelectedItem != null)
            {
                queryStringSQL = "SELECT * FROM " + viewListComboBox.SelectedItem.ToString();
            }
            else
            {
            }

            //criteria part of string
            string CriteriaString = this.ViewSQLString();

            if (CriteriaString != "")
            {
                queryStringSQL = queryStringSQL + " WHERE " + CriteriaString;
            }
            
            //get data from SQL DB
            DataTable myTmpData = ObjSQLData(queryStringSQL);
            
            //write to worksheet
            PastToWorksheet(myTmpData);
        }

        private void tableDownloadButton_Click(object sender, EventArgs e)
        {
            //SQL DB table download process

            //set all objects to null
            queryStringSQL = "";
            DataTable myData = new DataTable();

            //set SQL query
            if (tableListComboBox.SelectedItem != null)
            {
                queryStringSQL = "SELECT * FROM " + this.tableListComboBox.SelectedItem.ToString();
            }
            else
            {
            }

            //get data from SQL DB
            DataTable myTmpData = ObjSQLData(queryStringSQL);

            //write to worksheet
            PastToWorksheet(myTmpData);

            //Add columns
            Excel.Worksheet mySheet = Globals.ThisAddIn.Application.ActiveSheet;
            mySheet.Range["A1"].Offset[0, myTmpData.Columns.Count].Value = "Comm";
        }

        private void tableUploadbutton_Click(object sender, EventArgs e)
        {
            // Upload amended excel sheet to SQL table
            DataTable myUpTable = ObjSQLSchemaData("Columns", "ADS", this.tableListComboBox.SelectedItem.ToString());

            //query string to drop TestTab if exist
            string queryStringSQL = "IF OBJECT_ID('dbo.TestTab') IS NOT NULL DROP TABLE dbo.TestTab";
            queryStringSQL = queryStringSQL + " CREATE TABLE TestTab(";

            //create dynamic SQLstring
            var numRows = myUpTable.Rows.Count;
            for (var row = 0; row < numRows; row++)
            {
                queryStringSQL = queryStringSQL + myUpTable.Rows[row][3].ToString() + " " +
                    myUpTable.Rows[row][7].ToString();
                if (myUpTable.Rows[row][8].ToString() != "")
                {
                    queryStringSQL = queryStringSQL + "(" + myUpTable.Rows[row][8].ToString() + ")";
                }
                queryStringSQL = queryStringSQL + ",";
            }

            queryStringSQL = queryStringSQL +
                "[Comm] nvarchar(10))";
            
            
            //queryStringSQL = queryStringSQL +
            //    "[Command] varchar(10), [Result] varchar(10), [Comment] varchar(50))";

            //write datatable
            DataTable tmpUpdata = ExcelToDatatable();

            //upload
            using (SqlConnection connectionSQL = new SqlConnection(connectionStringSQL))
            using (SqlCommand querySQL = new SqlCommand(queryStringSQL, connectionSQL))
            using (SqlBulkCopy bulkCopySQL = new SqlBulkCopy(connectionSQL))
            {
                connectionSQL.Open();

                //delete TestTab
                querySQL.ExecuteNonQuery();

                //copy excel sheet to SQL table
                bulkCopySQL.DestinationTableName = "TestTab";
                bulkCopySQL.WriteToServer(tmpUpdata);

                
                //call stored procedure with input
                SqlCommand storedProcedureComm = new SqlCommand("sp_upd_0_UpdateLookups", connectionSQL);
                storedProcedureComm.CommandType = CommandType.StoredProcedure;
                storedProcedureComm.Parameters.Add(new SqlParameter("@TabInput", this.tableListComboBox.SelectedItem.ToString()));
                //storedProcedureComm.Parameters["@TabInput"].Value = this.comboBox1.SelectedItem.ToString();
                try
                {
                    storedProcedureComm.ExecuteNonQuery();
                    MessageBox.Show("Upload success");
                }
                catch
                {
                    MessageBox.Show("A problem with sp_upd_0_UpdateLookups.");
                }
                

                connectionSQL.Close();
            }
        }

        private void viewListComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            // update criterias to column name
            DataTable myView = ObjSQLSchemaData("Columns", "ADS", this.viewListComboBox.SelectedItem.ToString());

            this.criteriaComboBox1.Items.Clear();
            this.criteriaComboBox2.Items.Clear();
            this.criteriaComboBox3.Items.Clear();
            this.criteriaComboBox4.Items.Clear();
            this.criteriaComboBox5.Items.Clear();

            this.criteriaComboBox1.Items.Add("");
            this.criteriaComboBox2.Items.Add("");
            this.criteriaComboBox3.Items.Add("");
            this.criteriaComboBox4.Items.Add("");
            this.criteriaComboBox5.Items.Add("");

            var numRows = myView.Rows.Count;
            for (var row = 0; row < numRows; row++)
            {
                this.criteriaComboBox1.Items.Add(myView.Rows[row][3]);
                this.criteriaComboBox2.Items.Add(myView.Rows[row][3]);
                this.criteriaComboBox3.Items.Add(myView.Rows[row][3]);
                this.criteriaComboBox4.Items.Add(myView.Rows[row][3]);
                this.criteriaComboBox5.Items.Add(myView.Rows[row][3]);
            }

            this.criteriaComboBox1.SelectedItem = null;
            this.criteriaComboBox1.Text = "";
            this.criteriatextBox1.Text = "";

            this.criteriaComboBox2.SelectedItem = null;
            this.criteriaComboBox2.Text = "";
            this.criteriatextBox2.Text = "";

            this.criteriaComboBox3.SelectedItem = null;
            this.criteriaComboBox3.Text = "";
            this.criteriatextBox3.Text = "";

            this.criteriaComboBox4.SelectedItem = null;
            this.criteriaComboBox4.Text = "";
            this.criteriatextBox4.Text = "";

            this.criteriaComboBox5.SelectedItem = null;
            this.criteriaComboBox5.Text = "";
            this.criteriatextBox5.Text = "";
        }
    }
}

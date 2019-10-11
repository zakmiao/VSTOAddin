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
    class clsSQLModule
    {
        //Create SQL connection string
        string connectionStringSQL = @"Database=ADS;Server=CREREPSQL03;
                Integrated Security=True;connect timeout=60";

        //Create SQL quary string
        //string queryStringSQL;
        
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

        public void ObjSQLStoredProcedure(string StoredProcedure, string Parameter, string Guid)
        {
            //get SQL table as datatable

            using (SqlConnection connectionSQL = new SqlConnection(connectionStringSQL))
            using (SqlCommand querySQL = new SqlCommand(StoredProcedure, connectionSQL))
            {
                try
                {
                    querySQL.CommandType = CommandType.StoredProcedure;

                    Guid myGuid = new Guid(Guid);

                    querySQL.Parameters.Add(Parameter, SqlDbType.UniqueIdentifier).Value = myGuid;

                    connectionSQL.Open();
                    // run storedprocedure
                    querySQL.ExecuteNonQuery();

                    connectionSQL.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

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

        public void UploadToSQL(DataTable DataToUpload, String DestTable)
        {
            //query string to drop TestTab if exist
            string queryStringSQL = @"IF OBJECT_ID('dbo." + DestTable + @"') IS NOT NULL DROP TABLE dbo." + DestTable;

            queryStringSQL = queryStringSQL + @" CREATE TABLE " + DestTable + @"([SBFClass] nvarchar(50), "
            + @"[UWRef] nvarchar(50), [UWY] int, [EventCode] int, [DateMovement] date, [Reference] nvarchar(50), "
            + @"[LossNarr] nvarchar(250), [CCY] nchar(10), [Paid] float, [OS] float)";

            //upload
            using (SqlConnection connectionSQL = new SqlConnection(connectionStringSQL))
            using (SqlCommand querySQL = new SqlCommand(queryStringSQL, connectionSQL))
            using (SqlBulkCopy bulkCopySQL = new SqlBulkCopy(connectionSQL))
            {
                connectionSQL.Open();

                //delete TestTab
                querySQL.ExecuteNonQuery();

                //copy excel sheet to SQL table
                bulkCopySQL.DestinationTableName = DestTable;
                bulkCopySQL.WriteToServer(DataToUpload);
                
                /*
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
                */

                connectionSQL.Close();
            }
        }

        public void UploadMEToSQL(DataTable DataToUpload, String DestTable)
        {

            //query string to drop TestTab if exist
            string queryStringSQL = @"DELETE FROM " + DestTable;

            //upload
            using (SqlConnection connectionSQL = new SqlConnection(connectionStringSQL))
            using (SqlCommand querySQL = new SqlCommand(queryStringSQL, connectionSQL))
            using (SqlBulkCopy bulkCopySQL = new SqlBulkCopy(connectionSQL))
            {
                connectionSQL.Open();

                //delete table
                querySQL.ExecuteNonQuery();

                //copy excel sheet to SQL table
                bulkCopySQL.DestinationTableName = DestTable;
                bulkCopySQL.WriteToServer(DataToUpload);

                connectionSQL.Close();
            }
        }

        public void MESQLStoredProcedure(string StoredProcedure, string thisVersion, string thisAsAt)
        {
            //get SQL table as datatable

            using (SqlConnection connectionSQL = new SqlConnection(connectionStringSQL))
            using (SqlCommand querySQL = new SqlCommand(StoredProcedure, connectionSQL))
            {
                try
                {
                    querySQL.CommandType = CommandType.StoredProcedure;
                    
                    querySQL.Parameters.Add("Version", SqlDbType.NVarChar).Value = thisVersion;
                    querySQL.Parameters.Add("AsAt", SqlDbType.NVarChar).Value = thisAsAt;

                    connectionSQL.Open();
                    // run storedprocedure
                    querySQL.ExecuteNonQuery();

                    connectionSQL.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
    }
}

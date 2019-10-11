using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace ReservingUploader
{
    public partial class frmUploader_VersionControl : Form
    {
        public frmUploader_VersionControl()
        {
            InitializeComponent();
        }

        public void GetVersion(string AsAt)
        {
            //get SQL table as datatable
            DataTable myData = new DataTable();

            string connectionStringSQL = @"Database=ADS;Server=CREREPSQL03;
                Integrated Security=True;connect timeout=30";

            string SQLquery = @"SELECT DISTINCT [Version] FROM [lu_Version] ";

            SQLquery = SQLquery + @"WHERE [AsAt]='" + AsAt + @"'";
            
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

            var numRows = myData.Rows.Count;

            try
            {
                for (var row = 0; row < numRows; row++)
                    this.VersionComboBox.Items.Add(myData.Rows[row][0]);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public String Version
        {
            get { return VersionComboBox.Text; }
        }
    }
}

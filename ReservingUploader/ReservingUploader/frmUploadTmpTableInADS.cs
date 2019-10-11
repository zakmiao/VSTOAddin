using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ReservingUploader
{
    public partial class frmUploadTmpTableInADS : Form
    {
        //initial set up
        clsSQLModule mySQLModule = new clsSQLModule();

        public frmUploadTmpTableInADS()
        {
            InitializeComponent();
        }

        private void frmUploadTmpTableInADS_Load(object sender, EventArgs e)
        {
            
        }

        private void UploadButton_Click(object sender, EventArgs e)
        {
            try
            {
                mySQLModule.ObjSQLStoredProcedure(@"sp_imp_0_RunAllImport", @"@ImportID", this.comboBox1.SelectedItem.ToString());
                MessageBox.Show("Data Uploaded");
                this.Close();
            }
            catch
            {
                MessageBox.Show("Data not Uploaded");
            }
            
        }

        private void DeleteButton_Click(object sender, EventArgs e)
        {
            try
            {
                mySQLModule.ObjSQLStoredProcedure(@"sp_upd_2_DeleteVersion", @"@ImportID", this.comboBox1.SelectedItem.ToString());
                MessageBox.Show("Data deleted");
                this.Close();
            }
            catch
            {
                MessageBox.Show("Data not deleted");
            }
            
        }

        private void btn_CheckData_Click(object sender, EventArgs e)
        {
            try
            {
                DataTable DataPointNoMatch = mySQLModule.ObjSQLData(@"SELECT TOP(8) * FROM [vw_imp_1a_DataPointIDs_NoMatch]");
                this.dataGridView1.DataSource = DataPointNoMatch;
                this.dataGridView1.AutoResizeColumns();

                DataTable DataValueNoMatch = mySQLModule.ObjSQLData(@"SELECT TOP(8) * FROM [vw_imp_2a_DataValueIDs_NoMatch]");
                this.dataGridView2.DataSource = DataValueNoMatch;
                this.dataGridView2.AutoResizeColumns();

                DataTable GuidTable = mySQLModule.ObjSQLData(@"SELECT DISTINCT [ImportID] FROM [imp_tmp_UploadData]");
                for (int i = 0; i < GuidTable.Rows.Count; i++)
                {
                    this.comboBox1.Items.Add(GuidTable.Rows[i][0]);
                }

                MessageBox.Show("Data Checked");
            }
            catch
            {
                MessageBox.Show("Data Checking Error");
            }
        }
    }
}

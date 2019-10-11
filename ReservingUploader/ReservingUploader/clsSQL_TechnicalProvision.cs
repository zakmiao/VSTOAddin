using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using System.Windows.Forms;

namespace S2088ReservingTools
{
    class clsSQL_TechnicalProvision
    {
        //Create SQL connection string
        string connectionStringSQL = @"Database=SII_Model;Server=CREREPSQL03;
                Integrated Security=True;connect timeout=60";

        public void UploadClaimBand()
        {
            //upload claimBand tab to SQL Database
            Excel.Workbook tgtWkbk = Globals.ThisAddIn.Application.ActiveWorkbook;

            String TgtWkshtName = "Claim Band";

            Excel.Worksheet tgtWksht = tgtWkbk.Sheets[TgtWkshtName];

            Excel.Range myRange = (Excel.Range)tgtWksht.UsedRange;

            object[,] myData = myRange.Value2;

            Int32 intErow= myData.GetLength(0);

            Int32 intEcol = myData.GetLength(1);

            Int32 intSrow = 3;

            DataTable DataTable = new DataTable();

            DataTable.Columns.Add("Syndicate", typeof(int));
            DataTable.Columns.Add("Quarter", typeof(string));
            DataTable.Columns.Add("YoA", typeof(int));
            DataTable.Columns.Add("CoB", typeof(int));
            DataTable.Columns.Add("ReservingCurrency", typeof(string));
            DataTable.Columns.Add("Type", typeof(string));
            DataTable.Columns.Add("Signed Premium", typeof(double));
            DataTable.Columns.Add("SII Written Premium", typeof(double));
            DataTable.Columns.Add("Earned Premium", typeof(double));
            DataTable.Columns.Add("Paid Claims", typeof(double));
            DataTable.Columns.Add("Incurred Claims", typeof(double));
            DataTable.Columns.Add("UWY Ultimate Premium", typeof(double));
            DataTable.Columns.Add("SII Ultimate Claims", typeof(double));
            DataTable.Columns.Add("SII Earned Claims", typeof(double));
            DataTable.Columns.Add("SII Unearned Written Claims", typeof(double));
            DataTable.Columns.Add("SII Unwritten Claims", typeof(double));
            DataTable.Columns.Add("Earned Margin", typeof(double));
            DataTable.Columns.Add("Unearned Margin", typeof(double));
            DataTable.Columns.Add("GAAP Written Premium", typeof(double));
            DataTable.Columns.Add("GAAP Unearned Written Claims", typeof(double));
            DataTable.Columns.Add("GAAP Unwritten Claims", typeof(double));
            DataTable.Columns.Add("DAC", typeof(double));
            DataTable.Columns.Add("UPR", typeof(double));

            for (int tmpRow = intSrow; tmpRow <= intErow; tmpRow++)
                {

                    try
                    {
                        DataRow uploadRow = DataTable.NewRow();

                        uploadRow[0] = 2088;                                                //Syndicate
                        uploadRow[1] = "2019Q1";                                            //Current Quarter
                        if (myData[tmpRow, 2] != null)
                            uploadRow[2] = Convert.ToInt16(myData[tmpRow, 2].ToString());   //YOA
                        if (myData[tmpRow, 4] != null)
                            uploadRow[3] = Convert.ToInt16(myData[tmpRow, 4].ToString());   //Class of Business
                        if (myData[tmpRow, 3] != null)
                            uploadRow[4] = myData[tmpRow, 3].ToString();                    //Currency
                        if (myData[tmpRow, 5] != null)
                            uploadRow[5] = myData[tmpRow, 5].ToString();                    //Type
                        if (myData[tmpRow, 6] != null)
                            uploadRow[6] = Convert.ToDouble(myData[tmpRow, 6].ToString());  //Signed Premium
                        if (myData[tmpRow, 7] != null)
                            uploadRow[7] = Convert.ToDouble(myData[tmpRow, 7].ToString());  //SII Written Premium
                        if (myData[tmpRow, 8] != null)
                            uploadRow[8] = Convert.ToDouble(myData[tmpRow, 8].ToString());  //Earned Premium
                        if (myData[tmpRow, 9] != null)
                            uploadRow[9] = Convert.ToDouble(myData[tmpRow, 9].ToString());  //Paid Claims
                        if (myData[tmpRow, 10] != null)
                            uploadRow[10] = Convert.ToDouble(myData[tmpRow, 10].ToString());  //Incurred Claims
                        if (myData[tmpRow, 11] != null)
                            uploadRow[11] = Convert.ToDouble(myData[tmpRow, 11].ToString());  //UWY Ultimate Premium
                        if (myData[tmpRow, 12] != null)
                            uploadRow[12] = Convert.ToDouble(myData[tmpRow, 12].ToString());  //SII Ultimate Claims
                        if (myData[tmpRow, 13] != null)
                            uploadRow[13] = Convert.ToDouble(myData[tmpRow, 13].ToString());  //SII Earned Claims
                        if (myData[tmpRow, 14] != null)
                            uploadRow[14] = Convert.ToDouble(myData[tmpRow, 14].ToString());  //SII Unearned Written Claims
                        if (myData[tmpRow, 15] != null)
                            uploadRow[15] = Convert.ToDouble(myData[tmpRow, 15].ToString());  //SII Unwritten Claims
                        if (myData[tmpRow, 16] != null)
                            uploadRow[16] = Convert.ToDouble(myData[tmpRow, 16].ToString());  //Earned Margin
                        if (myData[tmpRow, 17] != null)
                            uploadRow[17] = Convert.ToDouble(myData[tmpRow, 17].ToString());  //Unearned Margin
                        if (myData[tmpRow, 18] != null)
                            uploadRow[18] = Convert.ToDouble(myData[tmpRow, 18].ToString());  //GAAP Written Premium
                        if (myData[tmpRow, 19] != null)
                            uploadRow[19] = Convert.ToDouble(myData[tmpRow, 19].ToString());  //GAAP Unearned Written Claims
                        if (myData[tmpRow, 20] != null)
                            uploadRow[20] = Convert.ToDouble(myData[tmpRow, 20].ToString());  //GAAP Unwritten Claims
                        if (myData[tmpRow, 21] != null)
                            uploadRow[21] = Convert.ToDouble(myData[tmpRow, 21].ToString());  //DAC
                        if (myData[tmpRow, 22] != null)
                            uploadRow[22] = Convert.ToDouble(myData[tmpRow, 22].ToString());  //UPR

                        DataTable.Rows.Add(uploadRow);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                }

            //upload
            using (SqlConnection connectionSQL = new SqlConnection(connectionStringSQL))
            using (SqlBulkCopy bulkCopySQL = new SqlBulkCopy(connectionSQL))
            {
                connectionSQL.Open();

                //copy excel sheet to SQL table
                bulkCopySQL.DestinationTableName = "tmp_TblClaimsBand";
                bulkCopySQL.WriteToServer(DataTable);

                connectionSQL.Close();
            }
        }

        public void tmpTest()
        {
            Excel.Application myApp = Globals.ThisAddIn.Application;

            MessageBox.Show(myApp.Visible.ToString());
        }

    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using System.Windows.Forms;
using System.Drawing;


namespace ReservingUploader
{
    public partial class ribbonS2088Reserving
    {
        private void ribbonS2088Reserving_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void ADSQuery_Click(object sender, RibbonControlEventArgs e)
        {   
            Globals.ThisAddIn.myCustomTaskPane.Visible = ((RibbonToggleButton)sender).Checked;
        }

        private void uploadToADSTmp_Click(object sender, RibbonControlEventArgs e)
        {
            frmUploader myUploader = new frmUploader();
            myUploader.StartPosition = FormStartPosition.CenterScreen;
            myUploader.Show();
        }

        private void uploadWithinADS_Click(object sender, RibbonControlEventArgs e)
        {
            frmUploadTmpTableInADS myTmpUploader = new frmUploadTmpTableInADS();
            myTmpUploader.StartPosition = FormStartPosition.CenterScreen;
            myTmpUploader.Show();
        }

        
        private void btn_tmp_Click(object sender, RibbonControlEventArgs e)
        {
            //open workbook, get data task
            clsExcelModule modExcel = new clsExcelModule();

            Excel.Application myExcel = Globals.ThisAddIn.Application;

            Excel.Workbook myWkbk = myExcel.Workbooks.Open(@"U:\Actuary\Reserving\2019\Q1\Data\Copy of February 2019 UPF - returned delinked.xlsx");
            
            DataTable myData = new DataTable();

            myData.Columns.Add("YOA", System.Type.GetType("System.Int32"));
            myData.Columns.Add("Metric", System.Type.GetType("System.String"));
            myData.Columns.Add("SBF", System.Type.GetType("System.String"));
            myData.Columns.Add("USD", System.Type.GetType("System.Double"));
            myData.Columns.Add("GBP", System.Type.GetType("System.Double"));
            myData.Columns.Add("CAD", System.Type.GetType("System.Double"));
            myData.Columns.Add("EUR", System.Type.GetType("System.Double"));
            myData.Columns.Add("AUD", System.Type.GetType("System.Double"));
            myData.Columns.Add("JPY", System.Type.GetType("System.Double"));

            foreach (Excel.Worksheet myWksht in myWkbk.Worksheets)
            {
                Excel.Range myRange = myWksht.UsedRange;

                object[,] XlData = myRange.Value2;

                var nColumn = myRange.Columns.Count;
                var nRow = myRange.Rows.Count;

                if (XlData[1, 2]?.ToString() == @"Syndicate 2088 Ultimate Premium Forecast")
                {
                    for (int row = 13; row <= nRow; row++)
                    {
                        if((Convert.ToDouble(XlData[row, 25]?.ToString())+ Convert.ToDouble(XlData[row, 22]?.ToString()) != 0) 
                            && (new[] {"GBP", "USD", "CAD", "EUR", "AUD", "JPY"}.Contains(XlData[row, 3]?.ToString())))
                        {
                            DataRow myRow = myData.NewRow();

                            myRow["YOA"] = Convert.ToInt32(XlData[row, 2]?.ToString());
                            myRow["Metric"] = @"Reinstatement";
                            myRow["SBF"] = myWksht.Name.ToString();
                            
                            if (XlData[row, 3]?.ToString() == "USD")
                            {
                                myRow["USD"] = Convert.ToDouble(XlData[row, 25]?.ToString()) + Convert.ToDouble(XlData[row, 22]?.ToString());
                            }

                            if (XlData[row, 3]?.ToString() == "GBP")
                            {
                                myRow["GBP"] = Convert.ToDouble(XlData[row, 25]?.ToString()) + Convert.ToDouble(XlData[row, 22]?.ToString());
                            }

                            if (XlData[row, 3]?.ToString() == "CAD")
                            {
                                myRow["CAD"] = Convert.ToDouble(XlData[row, 25]?.ToString()) + Convert.ToDouble(XlData[row, 22]?.ToString());
                            }

                            if (XlData[row, 3]?.ToString() == "EUR")
                            {
                                myRow["EUR"] = Convert.ToDouble(XlData[row, 25]?.ToString()) + Convert.ToDouble(XlData[row, 22]?.ToString());
                            }

                            if (XlData[row, 3]?.ToString() == "AUD")
                            {
                                myRow["AUD"] = Convert.ToDouble(XlData[row, 25]?.ToString()) + Convert.ToDouble(XlData[row, 22]?.ToString());
                            }

                            if (XlData[row, 3]?.ToString() == "JPY")
                            {
                                myRow["JPY"] = Convert.ToDouble(XlData[row, 25]?.ToString()) + Convert.ToDouble(XlData[row, 22]?.ToString());
                            }

                            myData.Rows.Add(myRow);
                        }
                    }
                }
            }

            MessageBox.Show(myData.Rows.Count.ToString());

            modExcel.PastToWorksheet(myData, myWkbk.Worksheets["tmpUlt"]);

        }

        
    }
}

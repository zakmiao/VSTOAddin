using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ErnstAndYoung.WinRes.Core;
using System.Windows.Forms;
using Excel=Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.Data;
using ErnstAndYoung.WinRes.Core.COM;

namespace S2088ReservingTools
{
    class clsWinResModule
    {
        clsExcelModule myExcelmod = new clsExcelModule();

        String[] WRDataType = new String[] {"Premium", "Premium", "Premium", "Premium", "Premium", "Premium",
                                            "Paid", "Paid", "Paid", "Paid", "Paid", "Paid",
                                            "OS", "OS", "OS", "OS", "OS", "OS"};

        String[] WROperator = new String[] {"1", "1", "1", "1", "1", "1",
                                            "-1", "-1", "-1", "-1", "-1", "-1",
                                            "-1", "-1", "-1", "-1", "-1", "-1"};

        String[] WRDesc = new String[] {"Premium_AUD", "Premium_CAD", "Premium_EUR", "Premium_GBP", "Premium_USD", "Premium_JPY",
                                        "Paid_AUD", "Paid_CAD", "Paid_EUR", "Paid_GBP", "Paid_USD", "Paid_JPY",
                                        "OS_AUD", "OS_CAD", "OS_EUR", "OS_GBP", "OS_USD", "OS_JPY" };

        String[] CCY = new String[] {"AUD", "CAD", "EUR", "GBP", "USD", "JPY",
                                     "AUD", "CAD", "EUR", "GBP", "USD", "JPY",
                                     "AUD", "CAD", "EUR", "GBP", "USD", "JPY"};

        String[] SQLDataType = new String[] {"Prem Amt", "Prem Amt", "Prem Amt", "Prem Amt", "Prem Amt", "Prem Amt",
                                             "Paid Amt", "Paid Amt", "Paid Amt", "Paid Amt", "Paid Amt", "Paid Amt",
                                             "OS", "OS", "OS", "OS", "OS", "OS"};
        
        String[] LRC = new String[] {"CA", "WB", "W", "TS", "T", "Q", "B", "XM", "W2", "XL", "XH", "CT", "FA", "AG", "D3",
                                     "D5", "E3", "E5", "E7", "E9", "EB", "EC", "GM", "GS", "GX", "KG", "KS", "MF", "NA",
                                     "NC", "NP", "RX", "PG", "SB", "6T", "CZ", "D9", "SC", "SO", "TO", "TR", "TU", "TW",
                                     "TX", "V", "VL", "VX", "W3", "W4", "XE", "XR", "WL", "X3", "XA", "XC", "XJ", "XT",
                                     "XU", "CB", "CC", "W6", "XF", "XG", "XN"};

        String[] reviewLRC = new String[] {"6T", "CC", "CZ", "D3", "D5", "D9", "E3", "E5", "E7",
                                            "E9", "EC", "FA", "GM", "GS", "GX", "KG", "KS", "NA",
                                            "NC", "NP", "PG", "SB", "SC", "SO", "TO", "TR", "TU",
                                            "TW", "TX", "V", "VL", "VX", "W3", "W4", "WL", "X3",
                                            "XA", "XC", "XE", "XJ", "XR", "XT", "XU", };


        String[] newLRC = new String[] { "P3", "EH", "LE", "B5", "P", "BB", "D8", "NL" };

        /*
        String[] LRC = new String[] {"TO", "TR", "TU", "TW",
                                     "TX", "V", "VL", "VX", "W3", "W4", "XE", "XR", "WL", "X3", "XA", "XC", "XJ", "XT",
                                     "XU", "CB", "CC", "W6", "XF", "XG", "XN"};
        */

        public void CreateFiles()
        {
            //Create WR files with Data from SQL and assumptions from last year

            WinRes WR = new WinRes();

            FileTriangleInfo fti;

            FileHeaderInfo fhi;

            WinResFileBuilder fb;

            DataTable TriangleData = new DataTable();

            Double[,] TriangleDouble;
            
            for (Int16 j = 0; j < 8; j++)
            {

                //fti = WR.Files.CreateTriangleInfo(1997, 1, 2019, 1, PeriodLength.Year, PeriodLength.Quarter, 1, 3, 2019, 3);
                fti = WR.Files.CreateTriangleInfo(1997, 1, 2018, 1, PeriodLength.Year, PeriodLength.Quarter, 1, 12, 2018, 3);

                //fhi = WR.Files.CreateHeaderInfo(fti, "LRC BM", LRC[j], "GBP", 1000, 1000, OriginPeriodType.UnderwritingPeriod);
                //fhi = WR.Files.CreateHeaderInfo(fti, "LRC BM", newLRC[j], "GBP", 1000, 1000, OriginPeriodType.UnderwritingPeriod);
                fhi = WR.Files.CreateHeaderInfo(fti, "LRC BM", "E9", "GBP", 1000, 1000, OriginPeriodType.UnderwritingPeriod);

                fb = WR.Files.CreateFileBuilder(fhi);

                for (Int16 i = 0; i < 18; i++)
                {
                    //TriangleData = GetTriangle(CCY[i], LRC[j], SQLDataType[i]);
                    TriangleData = GetTriangle(CCY[i], newLRC[j], SQLDataType[i]);

                    TriangleDouble = ReadTriangleData(TriangleData, fti.OriginPeriods.Count, fti.DevPeriods.Count);

                    fb.ComponentTriangles.AddByName(WRDataType[i], WROperator[i], WRDesc[i], ref TriangleDouble, true);

                    TriangleDouble = null;

                    TriangleData = null;
                }

                WinResFile newFile;

                //newFile = WR.Files.CreateFileFromFileBuilderWithAssumptions(fb, @"U:\Reserving\LRC\2018 WRs\" + LRC[j] + @".pjx");

                newFile = WR.Files.CreateFileFromFileBuilder(fb);

                //newFile.SaveAs(@"U:\Reserving\LRC\" + LRC[j] + @".pjx", true);
                newFile.SaveAs(@"U:\Actuary\Planning\2019\LRC\" + newLRC[j] + @".pjx", true);

                fb = null;

                fhi = null;

                fti = null;
            }
        }

        public Double[,] ReadTriangleData(DataTable TriangleData, int rowCount, int colCount)
        {
            
            var result = new Double[rowCount, colCount];

            for(int counter = 0; counter < TriangleData.Rows.Count; counter++)
            {
                result[TriangleData.Rows[counter].Field<int>(0), TriangleData.Rows[counter].Field<int>(1)] = TriangleData.Rows[counter].Field<Double>(2);
            }

            return result;
        }
        
        public void tmpTest2()
        {
            //Create WR files with Data from SQL and assumptions from last year

            WinRes WR = new WinRes();

            FileTriangleInfo fti;

            FileHeaderInfo fhi;

            WinResFileBuilder fb;

            DataTable TriangleData = new DataTable();

            Double[,] TriangleDouble;



            //fti = WR.Files.CreateTriangleInfo(1997, 1, 2019, 1, PeriodLength.Year, PeriodLength.Quarter, 1, 3, 2019, 3);
            //fti = WR.Files.CreateTriangleInfo(1997, 1, 2018, 1, PeriodLength.Year, PeriodLength.Quarter, 1, 12, 2018, 3);
            fti = WR.Files.CreateTriangleInfo(1998, 1, 2019, 1, PeriodLength.Year, PeriodLength.Quarter, 1, 3, 2019,3);

            //fhi = WR.Files.CreateHeaderInfo(fti, "LRC BM", LRC[j], "GBP", 1000, 1000, OriginPeriodType.UnderwritingPeriod);
            //fhi = WR.Files.CreateHeaderInfo(fti, "LRC BM", newLRC[j], "GBP", 1000, 1000, OriginPeriodType.UnderwritingPeriod);
            fhi = WR.Files.CreateHeaderInfo(fti, "LRC BM", "E9", "GBP", 1000, 1000, OriginPeriodType.UnderwritingPeriod);

                fb = WR.Files.CreateFileBuilder(fhi);

                for (Int16 i = 0; i < 18; i++)
                {
                    //TriangleData = GetTriangle(CCY[i], LRC[j], SQLDataType[i]);
                    TriangleData = GetTriangle(CCY[i], "E9", SQLDataType[i]);

                    TriangleDouble = ReadTriangleData(TriangleData, fti.OriginPeriods.Count, fti.DevPeriods.Count);

                    fb.ComponentTriangles.AddByName(WRDataType[i], WROperator[i], WRDesc[i], ref TriangleDouble, true);

                    TriangleDouble = null;

                    TriangleData = null;
                }

                WinResFile newFile;

                newFile = WR.Files.CreateFileFromFileBuilderWithAssumptions(fb, @"U:\Actuary\Planning\2019\LRC\AsAtQ42018\review LRC 2019\E9.pjx");

                //newFile = WR.Files.CreateFileFromFileBuilder(fb);

                //newFile.SaveAs(@"U:\Reserving\LRC\" + LRC[j] + @".pjx", true);
                newFile.SaveAs(@"U:\Actuary\Planning\2019\LRC\E9.pjx", true);

                fb = null;

                fhi = null;

                fti = null;

        }

        public void tmpTest()
        {
            
            DataTable OutPut;

            OutPut = GetTriangle("GBP", "E9", "Paid Amt");

            Double[,] result;

            result = ReadTriangleData(OutPut,23,89);

            //write array to worksheet
            Excel.Application myApp = Globals.ThisAddIn.Application;
            Excel.Workbook myWkbk = myApp.ActiveWorkbook;
            Excel.Worksheet mySheet = myWkbk.ActiveSheet;

            mySheet.Cells.Clear();
            try { mySheet.Range["A1"].Resize[23, 89].Value = result; }
            catch (Exception ex) { MessageBox.Show(ex.Message); }

            //myExcelmod.PastToWorksheet(OutPut);
            
        }

        public DataTable GetTriangle(String CCY, String LRC, String DataType)
        {
            //Return DataTable of Triangle from SQL

            String connectionStringSQL = @"Database=LRC_DB;Server=CREREPSQL03;
                Integrated Security=True;connect timeout=60";

            String StoredProcedure= @"[sp_out_LRC_1_CnvTriangle]";

            DataTable myData=new DataTable();

            using (SqlConnection connectionSQL = new SqlConnection(connectionStringSQL))
            using (SqlCommand querySQL = new SqlCommand(StoredProcedure, connectionSQL))
            {
                try
                {
                    querySQL.CommandType = CommandType.StoredProcedure;

                    querySQL.Parameters.Add("@inp_WRConfigID",SqlDbType.Int).Value = 1;
                    querySQL.Parameters.Add("@inp_CCY", SqlDbType.NVarChar).Value = CCY;
                    querySQL.Parameters.Add("@inp_LRC", SqlDbType.NVarChar).Value = LRC;
                    querySQL.Parameters.Add("@inp_DataType", SqlDbType.NVarChar).Value = DataType;

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

        public void RefreshWR(String LRC)
        {
            WinRes WR = new WinRes();
            
            WinResFile WRFile = WR.Files.Open(@"U:\Actuary\Planning\2019\LRC\new LRC 2019\" + LRC + @".pjx");
            //WinResFile WRFile = WR.Files.Open(@"U:\Actuary\Planning\2019\LRC\review LRC 2019\" + LRC + @".pjx");

            SelectedUltimate sel = null;

            //write array to worksheet
            Excel.Application myApp = Globals.ThisAddIn.Application;
            Excel.Workbook myWkbk = myApp.ActiveWorkbook;
            Excel.Worksheet templateWksht = myWkbk.Sheets["Template"];

            int WkshtIndex = templateWksht.Index;

            MessageBox.Show(WkshtIndex.ToString());

            templateWksht.Copy(templateWksht);

            Excel.Worksheet mySheet = myWkbk.Worksheets[WkshtIndex];

            mySheet.Name = LRC;

            mySheet.Range["B1"].Value2 = LRC;

            Excel.Range myRng = mySheet.Range["B6"];

            WRFile.SelectedUltimates.TryGet("Premium", out sel);
            
            for(Int16 intUWY = 0; intUWY < 22; intUWY++)
            {
                myRng.Offset[intUWY, 0].Value = sel[intUWY];
            }

            sel = null;

            WRFile.SelectedUltimates.TryGet("Claim", out sel);

            for (Int16 intUWY = 0; intUWY < 22; intUWY++)
            {
                myRng.Offset[intUWY, 2].Value = sel[intUWY];
            }

            sel = null;
            
            for (Int16 intUWY = 0; intUWY < 22; intUWY++)
            {
                myRng.Offset[intUWY, 1].Value = WRFile.AnalysedTriangles["Incurred"].LeadingDiagonal[intUWY];
            }

            WR.Files.Remove(WRFile);

            WR = null;
        }


        public void ImportChainLadderProfiles()
        {

            WinRes WR = new WinRes();

            WinResFile WRFile = WR.Files.Open(@"U:\Actuary\Planning\2018\Analysis\WinRes Files\LRC BM\Finalized Models\6T.pjx");

            ChainLadderMethod CLMethod;

            IProjectionMethod projMethod = null;

            Double[] DevFactors;

            Excel.Application myApp = Globals.ThisAddIn.Application;

            Excel.Workbook myWkbk = myApp.ActiveWorkbook;

            Excel.Worksheet myWksht = myWkbk.ActiveSheet;

            Excel.Range myRng = myWksht.Range["A1"];


            projMethod = WRFile.ProjectionMethods["Premiumcl"];

            CLMethod = (ChainLadderMethod)projMethod;

            DevFactors = CLMethod.SubModels[CLMethod.GetSelectedSubModelIndex(0)].GetFinalSelectedDevFactors();


            Int32 numColumns = CLMethod.SubModels[CLMethod.GetSelectedSubModelIndex(0)].FinalFactorProviders[0].LastDevFactorIndex + 1;


            myRng.Resize[1, numColumns].Value = DevFactors;

            //String results = string.Join(",", DevFactors);

            //MessageBox.Show(results);
        }

    }

}

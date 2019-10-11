using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace ReservingUploader
{
    public partial class ThisAddIn
    {
        //ADS Query task pane
        public Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //For ADS Query task pane
            Application.WindowActivate += Application_WindowActivate;
        }

        private void Application_WindowActivate(Excel.Workbook Wb, Excel.Window Wn)
        {
            // load ADS query task pane control
            //throw new NotImplementedException();

            myCustomTaskPane = this.CustomTaskPanes.Add(new crlADSQuery(), "SQL Database Functions");
            myCustomTaskPane.Width = 360;

            Globals.Ribbons.ribbonS2088Reserving.ADSQuery.Checked = myCustomTaskPane.Visible;
            myCustomTaskPane.VisibleChanged += MyCustomTaskPane_VisibleChanged;
        }

        private void MyCustomTaskPane_VisibleChanged(object sender, EventArgs e)
        {
            //throw new NotImplementedException();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}

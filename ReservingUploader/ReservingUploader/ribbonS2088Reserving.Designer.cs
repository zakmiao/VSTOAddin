namespace ReservingUploader
{
    partial class ribbonS2088Reserving : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ribbonS2088Reserving()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.ReservingTools = this.Factory.CreateRibbonGroup();
            this.ADSQuery = this.Factory.CreateRibbonToggleButton();
            this.uploadToADSTmp = this.Factory.CreateRibbonButton();
            this.uploadWithinADS = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.ReservingTools.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.ReservingTools);
            this.tab1.Label = "S2088 Reserving Tools";
            this.tab1.Name = "tab1";
            // 
            // ReservingTools
            // 
            this.ReservingTools.Items.Add(this.ADSQuery);
            this.ReservingTools.Items.Add(this.uploadToADSTmp);
            this.ReservingTools.Items.Add(this.uploadWithinADS);
            this.ReservingTools.Label = "Reserving Tools";
            this.ReservingTools.Name = "ReservingTools";
            // 
            // ADSQuery
            // 
            this.ADSQuery.Label = "ADS Query";
            this.ADSQuery.Name = "ADSQuery";
            this.ADSQuery.OfficeImageId = "ResultsPaneStartFindAndReplace";
            this.ADSQuery.ShowImage = true;
            this.ADSQuery.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ADSQuery_Click);
            // 
            // uploadToADSTmp
            // 
            this.uploadToADSTmp.Label = "Upload to ADS tmp table";
            this.uploadToADSTmp.Name = "uploadToADSTmp";
            this.uploadToADSTmp.OfficeImageId = "OutlineDemote";
            this.uploadToADSTmp.ShowImage = true;
            this.uploadToADSTmp.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.uploadToADSTmp_Click);
            // 
            // uploadWithinADS
            // 
            this.uploadWithinADS.Label = "Upload tmp table in ADS";
            this.uploadWithinADS.Name = "uploadWithinADS";
            this.uploadWithinADS.OfficeImageId = "OutlineDemoteToBodyText";
            this.uploadWithinADS.ShowImage = true;
            this.uploadWithinADS.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.uploadWithinADS_Click);
            // 
            // ribbonS2088Reserving
            // 
            this.Name = "ribbonS2088Reserving";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.ribbonS2088Reserving_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.ReservingTools.ResumeLayout(false);
            this.ReservingTools.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup ReservingTools;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton uploadToADSTmp;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton uploadWithinADS;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton ADSQuery;
    }

    partial class ThisRibbonCollection
    {
        internal ribbonS2088Reserving ribbonS2088Reserving
        {
            get { return this.GetRibbon<ribbonS2088Reserving>(); }
        }
    }
}

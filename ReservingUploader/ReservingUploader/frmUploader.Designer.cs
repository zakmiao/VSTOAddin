namespace ReservingUploader
{
    partial class frmUploader
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.FileDropDataGridView = new System.Windows.Forms.DataGridView();
            this.UploadButton = new System.Windows.Forms.Button();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripProgressBar1 = new System.Windows.Forms.ToolStripProgressBar();
            ((System.ComponentModel.ISupportInitialize)(this.FileDropDataGridView)).BeginInit();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // FileDropDataGridView
            // 
            this.FileDropDataGridView.AllowDrop = true;
            this.FileDropDataGridView.AllowUserToAddRows = false;
            this.FileDropDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.FileDropDataGridView.Location = new System.Drawing.Point(12, 12);
            this.FileDropDataGridView.Name = "FileDropDataGridView";
            this.FileDropDataGridView.Size = new System.Drawing.Size(774, 406);
            this.FileDropDataGridView.TabIndex = 2;
            this.FileDropDataGridView.RowHeaderMouseDoubleClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.FileDropDataGridView_RowHeaderMouseDoubleClick);
            this.FileDropDataGridView.DragDrop += new System.Windows.Forms.DragEventHandler(this.FileDropDataGridView_DragDrop);
            this.FileDropDataGridView.DragEnter += new System.Windows.Forms.DragEventHandler(this.FileDropDataGridView_DragEnter);
            // 
            // UploadButton
            // 
            this.UploadButton.Location = new System.Drawing.Point(330, 424);
            this.UploadButton.Name = "UploadButton";
            this.UploadButton.Size = new System.Drawing.Size(135, 23);
            this.UploadButton.TabIndex = 3;
            this.UploadButton.Text = "Upload";
            this.UploadButton.UseVisualStyleBackColor = true;
            this.UploadButton.Click += new System.EventHandler(this.UploadButton_Click);
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel1,
            this.toolStripProgressBar1});
            this.statusStrip1.Location = new System.Drawing.Point(0, 454);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(796, 22);
            this.statusStrip1.TabIndex = 4;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Size = new System.Drawing.Size(118, 17);
            this.toolStripStatusLabel1.Text = "toolStripStatusLabel1";
            // 
            // toolStripProgressBar1
            // 
            this.toolStripProgressBar1.Name = "toolStripProgressBar1";
            this.toolStripProgressBar1.Size = new System.Drawing.Size(100, 16);
            this.toolStripProgressBar1.Step = 1;
            // 
            // frmUploader
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(796, 476);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.UploadButton);
            this.Controls.Add(this.FileDropDataGridView);
            this.Name = "frmUploader";
            this.Text = "Upload files to ADS";
            this.Load += new System.EventHandler(this.frmUploader_Load);
            ((System.ComponentModel.ISupportInitialize)(this.FileDropDataGridView)).EndInit();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView FileDropDataGridView;
        private System.Windows.Forms.Button UploadButton;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.ToolStripProgressBar toolStripProgressBar1;
    }
}
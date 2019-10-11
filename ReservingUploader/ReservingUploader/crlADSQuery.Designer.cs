namespace ReservingUploader
{
    partial class crlADSQuery
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(crlADSQuery));
            this.label1 = new System.Windows.Forms.Label();
            this.tableRefresh = new System.Windows.Forms.Button();
            this.tableListComboBox = new System.Windows.Forms.ComboBox();
            this.tableDownloadButton = new System.Windows.Forms.Button();
            this.tableUploadbutton = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.ViewRefresh = new System.Windows.Forms.Button();
            this.viewListComboBox = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.criteriaComboBox1 = new System.Windows.Forms.ComboBox();
            this.criteriatextBox1 = new System.Windows.Forms.TextBox();
            this.criteriaComboBox2 = new System.Windows.Forms.ComboBox();
            this.criteriatextBox2 = new System.Windows.Forms.TextBox();
            this.criteriaComboBox3 = new System.Windows.Forms.ComboBox();
            this.criteriatextBox3 = new System.Windows.Forms.TextBox();
            this.viewDownload = new System.Windows.Forms.Button();
            this.criteriatextBox4 = new System.Windows.Forms.TextBox();
            this.criteriaComboBox4 = new System.Windows.Forms.ComboBox();
            this.criteriatextBox5 = new System.Windows.Forms.TextBox();
            this.criteriaComboBox5 = new System.Windows.Forms.ComboBox();
            this.UserControlTab = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.UserControlTab.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(2, 10);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(138, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Choose the table to update:";
            // 
            // tableRefresh
            // 
            this.tableRefresh.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.tableRefresh.FlatAppearance.BorderSize = 0;
            this.tableRefresh.Image = ((System.Drawing.Image)(resources.GetObject("tableRefresh.Image")));
            this.tableRefresh.Location = new System.Drawing.Point(316, 7);
            this.tableRefresh.Name = "tableRefresh";
            this.tableRefresh.Size = new System.Drawing.Size(18, 18);
            this.tableRefresh.TabIndex = 5;
            this.tableRefresh.UseVisualStyleBackColor = true;
            this.tableRefresh.Click += new System.EventHandler(this.tableRefresh_Click);
            // 
            // tableListComboBox
            // 
            this.tableListComboBox.DropDownHeight = 280;
            this.tableListComboBox.DropDownWidth = 180;
            this.tableListComboBox.FormattingEnabled = true;
            this.tableListComboBox.IntegralHeight = false;
            this.tableListComboBox.Location = new System.Drawing.Point(2, 30);
            this.tableListComboBox.Name = "tableListComboBox";
            this.tableListComboBox.Size = new System.Drawing.Size(332, 21);
            this.tableListComboBox.TabIndex = 6;
            // 
            // tableDownloadButton
            // 
            this.tableDownloadButton.Location = new System.Drawing.Point(2, 60);
            this.tableDownloadButton.Name = "tableDownloadButton";
            this.tableDownloadButton.Size = new System.Drawing.Size(80, 20);
            this.tableDownloadButton.TabIndex = 7;
            this.tableDownloadButton.Text = "Download";
            this.tableDownloadButton.UseVisualStyleBackColor = true;
            this.tableDownloadButton.Click += new System.EventHandler(this.tableDownloadButton_Click);
            // 
            // tableUploadbutton
            // 
            this.tableUploadbutton.Location = new System.Drawing.Point(254, 60);
            this.tableUploadbutton.Name = "tableUploadbutton";
            this.tableUploadbutton.Size = new System.Drawing.Size(80, 20);
            this.tableUploadbutton.TabIndex = 8;
            this.tableUploadbutton.Text = "Upload";
            this.tableUploadbutton.UseVisualStyleBackColor = true;
            this.tableUploadbutton.Click += new System.EventHandler(this.tableUploadbutton_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(2, 10);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(150, 13);
            this.label2.TabIndex = 9;
            this.label2.Text = "Choose the view to download:";
            // 
            // ViewRefresh
            // 
            this.ViewRefresh.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.ViewRefresh.FlatAppearance.BorderSize = 0;
            this.ViewRefresh.Image = ((System.Drawing.Image)(resources.GetObject("ViewRefresh.Image")));
            this.ViewRefresh.Location = new System.Drawing.Point(316, 7);
            this.ViewRefresh.Name = "ViewRefresh";
            this.ViewRefresh.Size = new System.Drawing.Size(18, 18);
            this.ViewRefresh.TabIndex = 10;
            this.ViewRefresh.UseVisualStyleBackColor = true;
            this.ViewRefresh.Click += new System.EventHandler(this.ViewRefresh_Click);
            // 
            // viewListComboBox
            // 
            this.viewListComboBox.DropDownHeight = 280;
            this.viewListComboBox.DropDownWidth = 180;
            this.viewListComboBox.FormattingEnabled = true;
            this.viewListComboBox.IntegralHeight = false;
            this.viewListComboBox.Location = new System.Drawing.Point(2, 30);
            this.viewListComboBox.Name = "viewListComboBox";
            this.viewListComboBox.Size = new System.Drawing.Size(332, 21);
            this.viewListComboBox.TabIndex = 11;
            this.viewListComboBox.SelectedIndexChanged += new System.EventHandler(this.viewListComboBox_SelectedIndexChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(2, 55);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(84, 13);
            this.label3.TabIndex = 12;
            this.label3.Text = "Edit filter criteria:";
            // 
            // criteriaComboBox1
            // 
            this.criteriaComboBox1.DropDownHeight = 280;
            this.criteriaComboBox1.DropDownWidth = 180;
            this.criteriaComboBox1.FormattingEnabled = true;
            this.criteriaComboBox1.IntegralHeight = false;
            this.criteriaComboBox1.Location = new System.Drawing.Point(2, 80);
            this.criteriaComboBox1.Name = "criteriaComboBox1";
            this.criteriaComboBox1.Size = new System.Drawing.Size(85, 21);
            this.criteriaComboBox1.TabIndex = 13;
            // 
            // criteriatextBox1
            // 
            this.criteriatextBox1.Location = new System.Drawing.Point(93, 81);
            this.criteriatextBox1.Name = "criteriatextBox1";
            this.criteriatextBox1.Size = new System.Drawing.Size(241, 20);
            this.criteriatextBox1.TabIndex = 14;
            // 
            // criteriaComboBox2
            // 
            this.criteriaComboBox2.DropDownHeight = 280;
            this.criteriaComboBox2.DropDownWidth = 180;
            this.criteriaComboBox2.FormattingEnabled = true;
            this.criteriaComboBox2.IntegralHeight = false;
            this.criteriaComboBox2.Location = new System.Drawing.Point(2, 107);
            this.criteriaComboBox2.Name = "criteriaComboBox2";
            this.criteriaComboBox2.Size = new System.Drawing.Size(85, 21);
            this.criteriaComboBox2.TabIndex = 15;
            // 
            // criteriatextBox2
            // 
            this.criteriatextBox2.Location = new System.Drawing.Point(93, 107);
            this.criteriatextBox2.Name = "criteriatextBox2";
            this.criteriatextBox2.Size = new System.Drawing.Size(241, 20);
            this.criteriatextBox2.TabIndex = 16;
            // 
            // criteriaComboBox3
            // 
            this.criteriaComboBox3.DropDownHeight = 280;
            this.criteriaComboBox3.DropDownWidth = 180;
            this.criteriaComboBox3.FormattingEnabled = true;
            this.criteriaComboBox3.IntegralHeight = false;
            this.criteriaComboBox3.Location = new System.Drawing.Point(2, 134);
            this.criteriaComboBox3.Name = "criteriaComboBox3";
            this.criteriaComboBox3.Size = new System.Drawing.Size(85, 21);
            this.criteriaComboBox3.TabIndex = 17;
            // 
            // criteriatextBox3
            // 
            this.criteriatextBox3.Location = new System.Drawing.Point(93, 135);
            this.criteriatextBox3.Name = "criteriatextBox3";
            this.criteriatextBox3.Size = new System.Drawing.Size(241, 20);
            this.criteriatextBox3.TabIndex = 18;
            // 
            // viewDownload
            // 
            this.viewDownload.Location = new System.Drawing.Point(2, 215);
            this.viewDownload.Name = "viewDownload";
            this.viewDownload.Size = new System.Drawing.Size(80, 20);
            this.viewDownload.TabIndex = 19;
            this.viewDownload.Text = "Download";
            this.viewDownload.UseVisualStyleBackColor = true;
            this.viewDownload.Click += new System.EventHandler(this.viewDownload_Click);
            // 
            // criteriatextBox4
            // 
            this.criteriatextBox4.Location = new System.Drawing.Point(93, 162);
            this.criteriatextBox4.Name = "criteriatextBox4";
            this.criteriatextBox4.Size = new System.Drawing.Size(241, 20);
            this.criteriatextBox4.TabIndex = 21;
            // 
            // criteriaComboBox4
            // 
            this.criteriaComboBox4.DropDownHeight = 280;
            this.criteriaComboBox4.DropDownWidth = 180;
            this.criteriaComboBox4.FormattingEnabled = true;
            this.criteriaComboBox4.IntegralHeight = false;
            this.criteriaComboBox4.Location = new System.Drawing.Point(2, 161);
            this.criteriaComboBox4.Name = "criteriaComboBox4";
            this.criteriaComboBox4.Size = new System.Drawing.Size(85, 21);
            this.criteriaComboBox4.TabIndex = 20;
            // 
            // criteriatextBox5
            // 
            this.criteriatextBox5.Location = new System.Drawing.Point(93, 189);
            this.criteriatextBox5.Name = "criteriatextBox5";
            this.criteriatextBox5.Size = new System.Drawing.Size(241, 20);
            this.criteriatextBox5.TabIndex = 23;
            // 
            // criteriaComboBox5
            // 
            this.criteriaComboBox5.DropDownHeight = 280;
            this.criteriaComboBox5.DropDownWidth = 180;
            this.criteriaComboBox5.FormattingEnabled = true;
            this.criteriaComboBox5.IntegralHeight = false;
            this.criteriaComboBox5.Location = new System.Drawing.Point(2, 188);
            this.criteriaComboBox5.Name = "criteriaComboBox5";
            this.criteriaComboBox5.Size = new System.Drawing.Size(85, 21);
            this.criteriaComboBox5.TabIndex = 22;
            // 
            // UserControlTab
            // 
            this.UserControlTab.Controls.Add(this.tabPage1);
            this.UserControlTab.Controls.Add(this.tabPage2);
            this.UserControlTab.Location = new System.Drawing.Point(0, 0);
            this.UserControlTab.Margin = new System.Windows.Forms.Padding(0);
            this.UserControlTab.Name = "UserControlTab";
            this.UserControlTab.SelectedIndex = 0;
            this.UserControlTab.Size = new System.Drawing.Size(344, 300);
            this.UserControlTab.TabIndex = 24;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.label1);
            this.tabPage1.Controls.Add(this.tableRefresh);
            this.tabPage1.Controls.Add(this.tableListComboBox);
            this.tabPage1.Controls.Add(this.tableDownloadButton);
            this.tabPage1.Controls.Add(this.tableUploadbutton);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(336, 274);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Update";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.label2);
            this.tabPage2.Controls.Add(this.criteriatextBox5);
            this.tabPage2.Controls.Add(this.ViewRefresh);
            this.tabPage2.Controls.Add(this.criteriaComboBox5);
            this.tabPage2.Controls.Add(this.viewListComboBox);
            this.tabPage2.Controls.Add(this.criteriatextBox4);
            this.tabPage2.Controls.Add(this.label3);
            this.tabPage2.Controls.Add(this.criteriaComboBox4);
            this.tabPage2.Controls.Add(this.criteriaComboBox1);
            this.tabPage2.Controls.Add(this.viewDownload);
            this.tabPage2.Controls.Add(this.criteriatextBox1);
            this.tabPage2.Controls.Add(this.criteriatextBox3);
            this.tabPage2.Controls.Add(this.criteriaComboBox2);
            this.tabPage2.Controls.Add(this.criteriaComboBox3);
            this.tabPage2.Controls.Add(this.criteriatextBox2);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(336, 274);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Download";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // crlADSQuery
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.UserControlTab);
            this.Name = "crlADSQuery";
            this.Size = new System.Drawing.Size(344, 377);
            this.UserControlTab.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button tableRefresh;
        private System.Windows.Forms.ComboBox tableListComboBox;
        private System.Windows.Forms.Button tableDownloadButton;
        private System.Windows.Forms.Button tableUploadbutton;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button ViewRefresh;
        private System.Windows.Forms.ComboBox viewListComboBox;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox criteriaComboBox1;
        private System.Windows.Forms.TextBox criteriatextBox1;
        private System.Windows.Forms.ComboBox criteriaComboBox2;
        private System.Windows.Forms.TextBox criteriatextBox2;
        private System.Windows.Forms.ComboBox criteriaComboBox3;
        private System.Windows.Forms.TextBox criteriatextBox3;
        private System.Windows.Forms.Button viewDownload;
        private System.Windows.Forms.TextBox criteriatextBox4;
        private System.Windows.Forms.ComboBox criteriaComboBox4;
        private System.Windows.Forms.TextBox criteriatextBox5;
        private System.Windows.Forms.ComboBox criteriaComboBox5;
        private System.Windows.Forms.TabControl UserControlTab;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
    }
}

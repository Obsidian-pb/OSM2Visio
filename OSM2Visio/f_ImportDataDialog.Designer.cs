namespace OSM2Visio
{
    partial class f_ImportDataDialog
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(f_ImportDataDialog));
            this.FD = new System.Windows.Forms.OpenFileDialog();
            this.TB_FilePath = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.B_SearchFile = new System.Windows.Forms.Button();
            this.CB_EWSSource = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.TB_EWSPath = new System.Windows.Forms.TextBox();
            this.B_SearchEWS = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.B_OK = new System.Windows.Forms.Button();
            this.B_Cancel = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // FD
            // 
            resources.ApplyResources(this.FD, "FD");
            // 
            // TB_FilePath
            // 
            resources.ApplyResources(this.TB_FilePath, "TB_FilePath");
            this.TB_FilePath.Name = "TB_FilePath";
            // 
            // label1
            // 
            resources.ApplyResources(this.label1, "label1");
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.Name = "label1";
            // 
            // B_SearchFile
            // 
            resources.ApplyResources(this.B_SearchFile, "B_SearchFile");
            this.B_SearchFile.Name = "B_SearchFile";
            this.B_SearchFile.UseVisualStyleBackColor = true;
            this.B_SearchFile.Click += new System.EventHandler(this.B_Search_Click);
            // 
            // CB_EWSSource
            // 
            this.CB_EWSSource.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.CB_EWSSource.FormattingEnabled = true;
            this.CB_EWSSource.Items.AddRange(new object[] {
            resources.GetString("CB_EWSSource.Items"),
            resources.GetString("CB_EWSSource.Items1"),
            resources.GetString("CB_EWSSource.Items2"),
            resources.GetString("CB_EWSSource.Items3")});
            resources.ApplyResources(this.CB_EWSSource, "CB_EWSSource");
            this.CB_EWSSource.Name = "CB_EWSSource";
            this.CB_EWSSource.SelectedIndexChanged += new System.EventHandler(this.CB_EWSSource_SelectedIndexChanged);
            // 
            // label2
            // 
            resources.ApplyResources(this.label2, "label2");
            this.label2.BackColor = System.Drawing.Color.Transparent;
            this.label2.Name = "label2";
            // 
            // TB_EWSPath
            // 
            resources.ApplyResources(this.TB_EWSPath, "TB_EWSPath");
            this.TB_EWSPath.Name = "TB_EWSPath";
            // 
            // B_SearchEWS
            // 
            resources.ApplyResources(this.B_SearchEWS, "B_SearchEWS");
            this.B_SearchEWS.Name = "B_SearchEWS";
            this.B_SearchEWS.UseVisualStyleBackColor = true;
            this.B_SearchEWS.Click += new System.EventHandler(this.B_SearchEWS_Click);
            // 
            // label3
            // 
            resources.ApplyResources(this.label3, "label3");
            this.label3.BackColor = System.Drawing.Color.Transparent;
            this.label3.Name = "label3";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.CB_EWSSource);
            this.groupBox1.Controls.Add(this.B_SearchEWS);
            this.groupBox1.Controls.Add(this.TB_EWSPath);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label3);
            resources.ApplyResources(this.groupBox1, "groupBox1");
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.TabStop = false;
            // 
            // linkLabel1
            // 
            resources.ApplyResources(this.linkLabel1, "linkLabel1");
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.TabStop = true;
            this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
            // 
            // groupBox2
            // 
            resources.ApplyResources(this.groupBox2, "groupBox2");
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.TabStop = false;
            // 
            // B_OK
            // 
            resources.ApplyResources(this.B_OK, "B_OK");
            this.B_OK.Name = "B_OK";
            this.B_OK.UseVisualStyleBackColor = true;
            this.B_OK.Click += new System.EventHandler(this.B_OK_Click);
            // 
            // B_Cancel
            // 
            resources.ApplyResources(this.B_Cancel, "B_Cancel");
            this.B_Cancel.Name = "B_Cancel";
            this.B_Cancel.UseVisualStyleBackColor = true;
            this.B_Cancel.Click += new System.EventHandler(this.B_Cancel_Click);
            // 
            // f_ImportDataDialog
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.B_Cancel);
            this.Controls.Add(this.B_OK);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.linkLabel1);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.B_SearchFile);
            this.Controls.Add(this.TB_FilePath);
            this.Controls.Add(this.label1);
            this.Name = "f_ImportDataDialog";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog FD;
        private System.Windows.Forms.TextBox TB_FilePath;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button B_SearchFile;
        private System.Windows.Forms.ComboBox CB_EWSSource;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox TB_EWSPath;
        private System.Windows.Forms.Button B_SearchEWS;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.LinkLabel linkLabel1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button B_OK;
        private System.Windows.Forms.Button B_Cancel;
    }
}
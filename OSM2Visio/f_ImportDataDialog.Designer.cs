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
            this.FD.Filter = "Файл данных OSM|*.osm";
            this.FD.Title = "Укажите расположение файла данных OSM";
            // 
            // TB_FilePath
            // 
            this.TB_FilePath.Location = new System.Drawing.Point(12, 27);
            this.TB_FilePath.Name = "TB_FilePath";
            this.TB_FilePath.Size = new System.Drawing.Size(620, 20);
            this.TB_FilePath.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.Location = new System.Drawing.Point(12, 11);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(103, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Файл данных OSM";
            // 
            // B_SearchFile
            // 
            this.B_SearchFile.Location = new System.Drawing.Point(638, 26);
            this.B_SearchFile.Name = "B_SearchFile";
            this.B_SearchFile.Size = new System.Drawing.Size(68, 20);
            this.B_SearchFile.TabIndex = 2;
            this.B_SearchFile.Text = "Обзор";
            this.B_SearchFile.UseVisualStyleBackColor = true;
            this.B_SearchFile.Click += new System.EventHandler(this.B_Search_Click);
            // 
            // CB_EWSSource
            // 
            this.CB_EWSSource.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.CB_EWSSource.FormattingEnabled = true;
            this.CB_EWSSource.Items.AddRange(new object[] {
            "Данные OSM (Карта 01)",
            "EWS (на базе Access)",
            "EWS (на базе SQL Server)",
            "ЭСУ ППВ"});
            this.CB_EWSSource.Location = new System.Drawing.Point(6, 32);
            this.CB_EWSSource.Name = "CB_EWSSource";
            this.CB_EWSSource.Size = new System.Drawing.Size(271, 21);
            this.CB_EWSSource.TabIndex = 3;
            this.CB_EWSSource.SelectedIndexChanged += new System.EventHandler(this.CB_EWSSource_SelectedIndexChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.Transparent;
            this.label2.Location = new System.Drawing.Point(6, 16);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(51, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Система";
            // 
            // TB_EWSPath
            // 
            this.TB_EWSPath.Enabled = false;
            this.TB_EWSPath.Location = new System.Drawing.Point(6, 72);
            this.TB_EWSPath.Name = "TB_EWSPath";
            this.TB_EWSPath.Size = new System.Drawing.Size(608, 20);
            this.TB_EWSPath.TabIndex = 0;
            // 
            // B_SearchEWS
            // 
            this.B_SearchEWS.Enabled = false;
            this.B_SearchEWS.Location = new System.Drawing.Point(620, 71);
            this.B_SearchEWS.Name = "B_SearchEWS";
            this.B_SearchEWS.Size = new System.Drawing.Size(68, 20);
            this.B_SearchEWS.TabIndex = 2;
            this.B_SearchEWS.Text = "Обзор";
            this.B_SearchEWS.UseVisualStyleBackColor = true;
            this.B_SearchEWS.Click += new System.EventHandler(this.B_SearchEWS_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.BackColor = System.Drawing.Color.Transparent;
            this.label3.Location = new System.Drawing.Point(6, 56);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(157, 13);
            this.label3.TabIndex = 1;
            this.label3.Text = "Расположение файла данных";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.CB_EWSSource);
            this.groupBox1.Controls.Add(this.B_SearchEWS);
            this.groupBox1.Controls.Add(this.TB_EWSPath);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Location = new System.Drawing.Point(12, 72);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(694, 104);
            this.groupBox1.TabIndex = 4;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Источник данных ИНППВ";
            // 
            // linkLabel1
            // 
            this.linkLabel1.AutoSize = true;
            this.linkLabel1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.linkLabel1.Location = new System.Drawing.Point(474, 50);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(158, 13);
            this.linkLabel1.TabIndex = 5;
            this.linkLabel1.TabStop = true;
            this.linkLabel1.Text = "Что такое Файл данных OSM?";
            this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
            // 
            // groupBox2
            // 
            this.groupBox2.Location = new System.Drawing.Point(12, 193);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(693, 231);
            this.groupBox2.TabIndex = 6;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Настройки";
            // 
            // B_OK
            // 
            this.B_OK.Location = new System.Drawing.Point(515, 430);
            this.B_OK.Name = "B_OK";
            this.B_OK.Size = new System.Drawing.Size(92, 26);
            this.B_OK.TabIndex = 7;
            this.B_OK.Text = "Готово";
            this.B_OK.UseVisualStyleBackColor = true;
            this.B_OK.Click += new System.EventHandler(this.B_OK_Click);
            // 
            // B_Cancel
            // 
            this.B_Cancel.Location = new System.Drawing.Point(613, 430);
            this.B_Cancel.Name = "B_Cancel";
            this.B_Cancel.Size = new System.Drawing.Size(92, 26);
            this.B_Cancel.TabIndex = 7;
            this.B_Cancel.Text = "Отмена";
            this.B_Cancel.UseVisualStyleBackColor = true;
            // 
            // f_ImportDataDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(717, 475);
            this.Controls.Add(this.B_Cancel);
            this.Controls.Add(this.B_OK);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.linkLabel1);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.B_SearchFile);
            this.Controls.Add(this.TB_FilePath);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "f_ImportDataDialog";
            this.Text = "Импорт картографических данных Open Street Map";
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
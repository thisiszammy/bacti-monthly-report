
namespace MonthlyReportEntrySoftware
{
    partial class FrmMain
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
            this.label1 = new System.Windows.Forms.Label();
            this.txtSelectedFilePath = new System.Windows.Forms.TextBox();
            this.btnSetFilePath = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.btnMonthlyReport = new System.Windows.Forms.Button();
            this.btnLoadFromFile = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(11, 26);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(90, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Select File Path : ";
            // 
            // txtSelectedFilePath
            // 
            this.txtSelectedFilePath.Enabled = false;
            this.txtSelectedFilePath.Location = new System.Drawing.Point(100, 23);
            this.txtSelectedFilePath.Name = "txtSelectedFilePath";
            this.txtSelectedFilePath.Size = new System.Drawing.Size(278, 20);
            this.txtSelectedFilePath.TabIndex = 3;
            // 
            // btnSetFilePath
            // 
            this.btnSetFilePath.Location = new System.Drawing.Point(382, 21);
            this.btnSetFilePath.Name = "btnSetFilePath";
            this.btnSetFilePath.Size = new System.Drawing.Size(32, 23);
            this.btnSetFilePath.TabIndex = 4;
            this.btnSetFilePath.Text = "...";
            this.btnSetFilePath.UseVisualStyleBackColor = true;
            this.btnSetFilePath.Click += new System.EventHandler(this.btnSetFilePath_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(22, 54);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(70, 13);
            this.label2.TabIndex = 6;
            this.label2.Text = "Report Forms";
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.btnLoadFromFile);
            this.panel1.Controls.Add(this.btnMonthlyReport);
            this.panel1.Location = new System.Drawing.Point(14, 61);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(400, 100);
            this.panel1.TabIndex = 7;
            // 
            // btnMonthlyReport
            // 
            this.btnMonthlyReport.Location = new System.Drawing.Point(11, 18);
            this.btnMonthlyReport.Name = "btnMonthlyReport";
            this.btnMonthlyReport.Size = new System.Drawing.Size(183, 62);
            this.btnMonthlyReport.TabIndex = 1;
            this.btnMonthlyReport.Text = "CREATE NEW MONTHLY REPORT";
            this.btnMonthlyReport.UseVisualStyleBackColor = true;
            this.btnMonthlyReport.Click += new System.EventHandler(this.btnMonthlyReport_Click);
            // 
            // btnLoadFromFile
            // 
            this.btnLoadFromFile.Location = new System.Drawing.Point(204, 18);
            this.btnLoadFromFile.Name = "btnLoadFromFile";
            this.btnLoadFromFile.Size = new System.Drawing.Size(183, 62);
            this.btnLoadFromFile.TabIndex = 2;
            this.btnLoadFromFile.Text = "LOAD MONTHLY REPORT FROM EXCEL FILE";
            this.btnLoadFromFile.UseVisualStyleBackColor = true;
            this.btnLoadFromFile.Click += new System.EventHandler(this.btnLoadFromFile_Click);
            // 
            // FrmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(423, 171);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.btnSetFilePath);
            this.Controls.Add(this.txtSelectedFilePath);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FrmMain";
            this.Text = "Report Form System";
            this.Load += new System.EventHandler(this.FrmMain_Load);
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtSelectedFilePath;
        private System.Windows.Forms.Button btnSetFilePath;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button btnMonthlyReport;
        private System.Windows.Forms.Button btnLoadFromFile;
    }
}


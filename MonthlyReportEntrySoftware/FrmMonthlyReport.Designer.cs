
namespace MonthlyReportEntrySoftware
{
    partial class FrmMonthlyReport
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
            this.panel1 = new System.Windows.Forms.Panel();
            this.txtMonth = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtYear = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.gridTestTypes = new System.Windows.Forms.DataGridView();
            this.txtTestType = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.IP = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ER = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.MAGS = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.OPD = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.txtSubTotal = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnSave = new System.Windows.Forms.Button();
            this.gridTallies = new System.Windows.Forms.DataGridView();
            this.txtSpecimen = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.txtIP = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.txtER = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.txtMAGS = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.txtOPD = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.txtTotal = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.label5 = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.panel3 = new System.Windows.Forms.Panel();
            this.btnBack = new System.Windows.Forms.Button();
            this.label6 = new System.Windows.Forms.Label();
            this.panel4 = new System.Windows.Forms.Panel();
            this.gridWeeks = new System.Windows.Forms.DataGridView();
            this.txtWeek = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.txtWIP = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.txtWER = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.txtWMAGS = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.txtWOPD = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.txtWTotal = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.label7 = new System.Windows.Forms.Label();
            this.txtFileName = new System.Windows.Forms.TextBox();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gridTestTypes)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridTallies)).BeginInit();
            this.panel2.SuspendLayout();
            this.panel3.SuspendLayout();
            this.panel4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gridWeeks)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.txtMonth);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.txtYear);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Location = new System.Drawing.Point(9, 45);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(249, 94);
            this.panel1.TabIndex = 0;
            // 
            // txtMonth
            // 
            this.txtMonth.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.txtMonth.FormattingEnabled = true;
            this.txtMonth.Location = new System.Drawing.Point(58, 52);
            this.txtMonth.Name = "txtMonth";
            this.txtMonth.Size = new System.Drawing.Size(169, 21);
            this.txtMonth.TabIndex = 3;
            this.txtMonth.SelectedValueChanged += new System.EventHandler(this.txtMonth_SelectedValueChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(15, 55);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(46, 13);
            this.label3.TabIndex = 2;
            this.label3.Text = "Month : ";
            // 
            // txtYear
            // 
            this.txtYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.txtYear.FormattingEnabled = true;
            this.txtYear.Location = new System.Drawing.Point(58, 22);
            this.txtYear.Name = "txtYear";
            this.txtYear.Size = new System.Drawing.Size(169, 21);
            this.txtYear.TabIndex = 1;
            this.txtYear.SelectedValueChanged += new System.EventHandler(this.txtYear_SelectedValueChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(23, 25);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(38, 13);
            this.label2.TabIndex = 0;
            this.label2.Text = "Year : ";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(17, 40);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(77, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Report Details ";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(275, 162);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(60, 13);
            this.label4.TabIndex = 3;
            this.label4.Text = "Test Types";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // gridTestTypes
            // 
            this.gridTestTypes.AllowUserToAddRows = false;
            this.gridTestTypes.AllowUserToDeleteRows = false;
            this.gridTestTypes.AllowUserToResizeColumns = false;
            this.gridTestTypes.AllowUserToResizeRows = false;
            this.gridTestTypes.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.gridTestTypes.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gridTestTypes.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.txtTestType,
            this.IP,
            this.ER,
            this.MAGS,
            this.OPD,
            this.txtSubTotal});
            this.gridTestTypes.Location = new System.Drawing.Point(13, 14);
            this.gridTestTypes.MultiSelect = false;
            this.gridTestTypes.Name = "gridTestTypes";
            this.gridTestTypes.ReadOnly = true;
            this.gridTestTypes.RowHeadersVisible = false;
            this.gridTestTypes.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            this.gridTestTypes.Size = new System.Drawing.Size(403, 378);
            this.gridTestTypes.TabIndex = 4;
            this.gridTestTypes.SelectionChanged += new System.EventHandler(this.gridTestTypes_SelectionChanged);
            // 
            // txtTestType
            // 
            this.txtTestType.FillWeight = 80F;
            this.txtTestType.HeaderText = "Type";
            this.txtTestType.Name = "txtTestType";
            this.txtTestType.ReadOnly = true;
            // 
            // IP
            // 
            this.IP.FillWeight = 30F;
            this.IP.HeaderText = "IP";
            this.IP.Name = "IP";
            this.IP.ReadOnly = true;
            // 
            // ER
            // 
            this.ER.FillWeight = 30F;
            this.ER.HeaderText = "ER";
            this.ER.Name = "ER";
            this.ER.ReadOnly = true;
            // 
            // MAGS
            // 
            this.MAGS.FillWeight = 30F;
            this.MAGS.HeaderText = "MAGS";
            this.MAGS.Name = "MAGS";
            this.MAGS.ReadOnly = true;
            // 
            // OPD
            // 
            this.OPD.FillWeight = 30F;
            this.OPD.HeaderText = "OPD";
            this.OPD.Name = "OPD";
            this.OPD.ReadOnly = true;
            // 
            // txtSubTotal
            // 
            this.txtSubTotal.FillWeight = 30F;
            this.txtSubTotal.HeaderText = "Total";
            this.txtSubTotal.Name = "txtSubTotal";
            this.txtSubTotal.ReadOnly = true;
            this.txtSubTotal.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.txtSubTotal.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(9, 145);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(249, 39);
            this.btnSave.TabIndex = 5;
            this.btnSave.Text = "Generate Report";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // gridTallies
            // 
            this.gridTallies.AllowUserToAddRows = false;
            this.gridTallies.AllowUserToDeleteRows = false;
            this.gridTallies.AllowUserToResizeColumns = false;
            this.gridTallies.AllowUserToResizeRows = false;
            this.gridTallies.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.gridTallies.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gridTallies.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.txtSpecimen,
            this.txtIP,
            this.txtER,
            this.txtMAGS,
            this.txtOPD,
            this.txtTotal});
            this.gridTallies.Location = new System.Drawing.Point(13, 14);
            this.gridTallies.MultiSelect = false;
            this.gridTallies.Name = "gridTallies";
            this.gridTallies.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            this.gridTallies.Size = new System.Drawing.Size(517, 531);
            this.gridTallies.TabIndex = 6;
            this.gridTallies.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.gridTallies_CellValueChanged);
            this.gridTallies.KeyDown += new System.Windows.Forms.KeyEventHandler(this.gridTallies_KeyDown);
            // 
            // txtSpecimen
            // 
            this.txtSpecimen.FillWeight = 80F;
            this.txtSpecimen.HeaderText = "Specimen";
            this.txtSpecimen.Name = "txtSpecimen";
            this.txtSpecimen.ReadOnly = true;
            // 
            // txtIP
            // 
            this.txtIP.FillWeight = 20F;
            this.txtIP.HeaderText = "IP";
            this.txtIP.Name = "txtIP";
            // 
            // txtER
            // 
            this.txtER.FillWeight = 20F;
            this.txtER.HeaderText = "ER";
            this.txtER.Name = "txtER";
            // 
            // txtMAGS
            // 
            this.txtMAGS.FillWeight = 20F;
            this.txtMAGS.HeaderText = "MAGS";
            this.txtMAGS.Name = "txtMAGS";
            // 
            // txtOPD
            // 
            this.txtOPD.FillWeight = 20F;
            this.txtOPD.HeaderText = "OPD";
            this.txtOPD.Name = "txtOPD";
            // 
            // txtTotal
            // 
            this.txtTotal.FillWeight = 20F;
            this.txtTotal.HeaderText = "Total";
            this.txtTotal.Name = "txtTotal";
            this.txtTotal.ReadOnly = true;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(711, 8);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(59, 13);
            this.label5.TabIndex = 7;
            this.label5.Text = "Specimens";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // panel2
            // 
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel2.Controls.Add(this.gridTallies);
            this.panel2.Location = new System.Drawing.Point(700, 14);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(546, 559);
            this.panel2.TabIndex = 8;
            // 
            // panel3
            // 
            this.panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel3.Controls.Add(this.gridTestTypes);
            this.panel3.Location = new System.Drawing.Point(264, 167);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(430, 405);
            this.panel3.TabIndex = 9;
            // 
            // btnBack
            // 
            this.btnBack.Location = new System.Drawing.Point(9, 533);
            this.btnBack.Name = "btnBack";
            this.btnBack.Size = new System.Drawing.Size(249, 39);
            this.btnBack.TabIndex = 10;
            this.btnBack.Text = "Back";
            this.btnBack.UseVisualStyleBackColor = true;
            this.btnBack.Click += new System.EventHandler(this.btnBack_Click);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(275, 9);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(44, 13);
            this.label6.TabIndex = 11;
            this.label6.Text = "Weeks ";
            this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // panel4
            // 
            this.panel4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel4.Controls.Add(this.gridWeeks);
            this.panel4.Location = new System.Drawing.Point(264, 14);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(430, 145);
            this.panel4.TabIndex = 12;
            // 
            // gridWeeks
            // 
            this.gridWeeks.AllowUserToAddRows = false;
            this.gridWeeks.AllowUserToDeleteRows = false;
            this.gridWeeks.AllowUserToResizeColumns = false;
            this.gridWeeks.AllowUserToResizeRows = false;
            this.gridWeeks.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.gridWeeks.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gridWeeks.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.txtWeek,
            this.txtWIP,
            this.txtWER,
            this.txtWMAGS,
            this.txtWOPD,
            this.txtWTotal});
            this.gridWeeks.Location = new System.Drawing.Point(13, 14);
            this.gridWeeks.MultiSelect = false;
            this.gridWeeks.Name = "gridWeeks";
            this.gridWeeks.ReadOnly = true;
            this.gridWeeks.RowHeadersVisible = false;
            this.gridWeeks.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            this.gridWeeks.Size = new System.Drawing.Size(403, 117);
            this.gridWeeks.TabIndex = 4;
            this.gridWeeks.SelectionChanged += new System.EventHandler(this.gridWeeks_SelectionChanged);
            // 
            // txtWeek
            // 
            this.txtWeek.FillWeight = 80F;
            this.txtWeek.HeaderText = "Week";
            this.txtWeek.Name = "txtWeek";
            this.txtWeek.ReadOnly = true;
            // 
            // txtWIP
            // 
            this.txtWIP.FillWeight = 30F;
            this.txtWIP.HeaderText = "IP";
            this.txtWIP.Name = "txtWIP";
            this.txtWIP.ReadOnly = true;
            // 
            // txtWER
            // 
            this.txtWER.FillWeight = 30F;
            this.txtWER.HeaderText = "ER";
            this.txtWER.Name = "txtWER";
            this.txtWER.ReadOnly = true;
            // 
            // txtWMAGS
            // 
            this.txtWMAGS.FillWeight = 30F;
            this.txtWMAGS.HeaderText = "MAGS";
            this.txtWMAGS.Name = "txtWMAGS";
            this.txtWMAGS.ReadOnly = true;
            // 
            // txtWOPD
            // 
            this.txtWOPD.FillWeight = 30F;
            this.txtWOPD.HeaderText = "OPD";
            this.txtWOPD.Name = "txtWOPD";
            this.txtWOPD.ReadOnly = true;
            // 
            // txtWTotal
            // 
            this.txtWTotal.FillWeight = 30F;
            this.txtWTotal.HeaderText = "Total";
            this.txtWTotal.Name = "txtWTotal";
            this.txtWTotal.ReadOnly = true;
            this.txtWTotal.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.txtWTotal.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(8, 17);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(63, 13);
            this.label7.TabIndex = 13;
            this.label7.Text = "File Name : ";
            // 
            // txtFileName
            // 
            this.txtFileName.Location = new System.Drawing.Point(68, 14);
            this.txtFileName.Name = "txtFileName";
            this.txtFileName.Size = new System.Drawing.Size(190, 20);
            this.txtFileName.TabIndex = 14;
            // 
            // FrmMonthlyReport
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1258, 581);
            this.Controls.Add(this.txtFileName);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.panel4);
            this.Controls.Add(this.btnBack);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.panel3);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.panel1);
            this.Name = "FrmMonthlyReport";
            this.Text = "Report Form System : Monthly Report";
            this.Load += new System.EventHandler(this.FrmMonthlyReport_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gridTestTypes)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridTallies)).EndInit();
            this.panel2.ResumeLayout(false);
            this.panel3.ResumeLayout(false);
            this.panel4.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.gridWeeks)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.ComboBox txtMonth;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox txtYear;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.DataGridView gridTestTypes;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.DataGridView gridTallies;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.DataGridViewTextBoxColumn txtSpecimen;
        private System.Windows.Forms.DataGridViewTextBoxColumn txtIP;
        private System.Windows.Forms.DataGridViewTextBoxColumn txtER;
        private System.Windows.Forms.DataGridViewTextBoxColumn txtMAGS;
        private System.Windows.Forms.DataGridViewTextBoxColumn txtOPD;
        private System.Windows.Forms.DataGridViewTextBoxColumn txtTotal;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Button btnBack;
        private System.Windows.Forms.DataGridViewTextBoxColumn txtTestType;
        private System.Windows.Forms.DataGridViewTextBoxColumn IP;
        private System.Windows.Forms.DataGridViewTextBoxColumn ER;
        private System.Windows.Forms.DataGridViewTextBoxColumn MAGS;
        private System.Windows.Forms.DataGridViewTextBoxColumn OPD;
        private System.Windows.Forms.DataGridViewTextBoxColumn txtSubTotal;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.DataGridView gridWeeks;
        private System.Windows.Forms.DataGridViewTextBoxColumn txtWeek;
        private System.Windows.Forms.DataGridViewTextBoxColumn txtWIP;
        private System.Windows.Forms.DataGridViewTextBoxColumn txtWER;
        private System.Windows.Forms.DataGridViewTextBoxColumn txtWMAGS;
        private System.Windows.Forms.DataGridViewTextBoxColumn txtWOPD;
        private System.Windows.Forms.DataGridViewTextBoxColumn txtWTotal;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtFileName;
    }
}
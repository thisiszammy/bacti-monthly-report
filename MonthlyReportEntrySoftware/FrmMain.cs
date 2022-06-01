using MonthlyReportEntrySoftware.Entities;
using MonthlyReportEntrySoftware.Services;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MonthlyReportEntrySoftware
{
    public partial class FrmMain : Form
    {
        private readonly ApplicationSettings applicationSettings;
        public FrmMain(ApplicationSettings applicationSettings)
        {

            this.StartPosition = FormStartPosition.CenterScreen;
            this.applicationSettings = applicationSettings;
            InitializeComponent();
        }

        private void FrmMain_Load(object sender, EventArgs e)
        {
            applicationSettings.AppForm = this;
        }

        private void btnSetFilePath_Click(object sender, EventArgs e)
        {
            using(var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Select Folder ...";
                dialog.ShowNewFolderButton = true;

                if(dialog.ShowDialog() == DialogResult.OK)
                {
                    applicationSettings.SavePath = dialog.SelectedPath;
                }
                else
                {
                    txtSelectedFilePath.Text = string.Empty;
                    applicationSettings.SavePath = null;
                }

                txtSelectedFilePath.Text = dialog.SelectedPath;
            }
        }


        private void btnMonthlyReport_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(applicationSettings.SavePath))
            {
                MessageBox.Show("Please Select Folder Path!");
                return;
            }

            FrmMonthlyReport frmMonthlyReport = new FrmMonthlyReport(applicationSettings, null);
            frmMonthlyReport.StartPosition = FormStartPosition.CenterScreen;
            frmMonthlyReport.Show();
            this.Hide();

        }

        private void btnLoadFromFile_Click(object sender, EventArgs e)
        {
            MonthlyReportEntity monthlyReportEntity = new MonthlyReportEntity();

            using(OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "C:\\";
                openFileDialog.Filter = "*.xls|*.xlsx";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if(openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    FileInfo fileInfo = new FileInfo(openFileDialog.FileName);
                    monthlyReportEntity.FileName = fileInfo.Name;
                    monthlyReportEntity.FilePath = fileInfo.DirectoryName;

                }

            }

            ReportGeneratorService.LoadFromExcelFile(monthlyReportEntity);

            FrmMonthlyReport frmMonthlyReport = new FrmMonthlyReport(applicationSettings, monthlyReportEntity);
            frmMonthlyReport.StartPosition = FormStartPosition.CenterScreen;
            frmMonthlyReport.Show();
            this.Hide();

        }
    }
}

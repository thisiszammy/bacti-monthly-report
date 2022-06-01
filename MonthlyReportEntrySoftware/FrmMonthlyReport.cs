using MonthlyReportEntrySoftware.Entities;
using MonthlyReportEntrySoftware.Entities.BaseEntities;
using MonthlyReportEntrySoftware.Services;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MonthlyReportEntrySoftware
{
    public partial class FrmMonthlyReport : Form
    {
        private readonly ApplicationSettings applicationSettings;
        private MonthlyReportEntity monthlyReportEntity;
        private List<CategoryTally> categoryTallies;
        private List<TestCategory> selectedWeekTestCategories;
        private string selectedTestType;
        private int selectedKey;
        public FrmMonthlyReport(ApplicationSettings applicationSettings, MonthlyReportEntity monthlyReportEntity)
        {
            this.applicationSettings = applicationSettings;
            this.monthlyReportEntity = (monthlyReportEntity == null) ? new MonthlyReportEntity() : monthlyReportEntity;
            InitializeComponent();
        }

        private void FrmMonthlyReport_Load(object sender, EventArgs e)
        {
            txtYear.SelectedValueChanged -= txtYear_SelectedValueChanged;
            txtMonth.SelectedValueChanged -= txtMonth_SelectedValueChanged;

            txtMonth.DataSource = DateTimeService.GetMonths().Select(x=> new { Key = x }).ToList();
            txtMonth.DisplayMember = "Key";
            txtMonth.ValueMember = "Key";

            txtYear.DataSource = DateTimeService.GenerateYearSpan(5).Select(x => new { Key = x }).ToList();
            txtYear.DisplayMember = "Key";
            txtYear.ValueMember = "Key";

            txtYear.SelectedValue = DateTime.Now.Year;
            txtMonth.SelectedValue = DateTimeService.GetMonthFromCode(DateTime.Now.Month);


            txtYear.SelectedIndexChanged += txtYear_SelectedValueChanged;
            txtMonth.SelectedIndexChanged += txtMonth_SelectedValueChanged;

            txtFileName.Text = monthlyReportEntity.FileName;

            LoadWeeks();
            gridWeeks.ClearSelection();
        }

        private void LoadWeeks()
        {
            gridWeeks.SelectionChanged -= gridWeeks_SelectionChanged;
            gridWeeks.Rows.Clear();

            foreach(var item in monthlyReportEntity.TestTallies)
            {
                selectedWeekTestCategories = (List<TestCategory>)item.Value;
                gridWeeks.Rows.Add(item.Key, selectedWeekTestCategories.Sum(x => 
                    x.CategoryTallies.Sum(y => y.IP)), 
                    selectedWeekTestCategories.Sum(x => x.CategoryTallies.Sum(y => y.ER)),
                    selectedWeekTestCategories.Sum(x => x.CategoryTallies.Sum(y => y.MAGS)), 
                    selectedWeekTestCategories.Sum(x => x.CategoryTallies.Sum(y => y.OPD)), 
                    selectedWeekTestCategories.Sum(x => x.CategoryTallyTotal));
            }


            gridWeeks.SelectionChanged += gridWeeks_SelectionChanged;
        }

        private void LoadTestTypes()
        {
            gridTestTypes.SelectionChanged -= gridTestTypes_SelectionChanged;
            gridTestTypes.Rows.Clear();

            foreach(var item in monthlyReportEntity.TestTallies[selectedKey])
            {
                gridTestTypes.Rows.Add(item.CategoryName, 
                    item.CategoryTallies.Sum(x=>x.IP), 
                    item.CategoryTallies.Sum(x => x.ER), 
                    item.CategoryTallies.Sum(x => x.MAGS), 
                    item.CategoryTallies.Sum(x => x.OPD), 
                    item.CategoryTallies.Sum(x => x.Total));
            }

            gridTestTypes.ClearSelection();
            gridTestTypes.SelectionChanged += gridTestTypes_SelectionChanged;
        }


        private void btnSave_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtFileName.Text.Trim()))
            {
                MessageBox.Show("Please Enter File Name!");
                return;
            }

            monthlyReportEntity.FilePath = applicationSettings.SavePath;
            monthlyReportEntity.FileName = txtFileName.Text.Trim();
            monthlyReportEntity.Year = (int)txtYear.SelectedValue;
            monthlyReportEntity.MonthCode = DateTimeService.GetMonthCodeFromString(txtMonth.SelectedValue.ToString());

            try
            {
                ReportGeneratorService.GenerateMonthlyReport(monthlyReportEntity);

                MessageBox.Show("Successfully Generated Report!");
            }
            catch (Exception ex)
            {
                if(ex.Message.Contains("being used by another process"))
                {
                    MessageBox.Show($"Please Close The Excel File `{monthlyReportEntity.FileName}`");
                }
            }
        }

        private void txtYear_SelectedValueChanged(object sender, EventArgs e)
        {

            monthlyReportEntity.Year = (int)txtYear.SelectedValue;
        }

        private void txtMonth_SelectedValueChanged(object sender, EventArgs e)
        {
            monthlyReportEntity.MonthCode = DateTimeService.GetMonthCodeFromString(txtMonth.SelectedValue.ToString());
        }

        private void gridTestTypes_SelectionChanged(object sender, EventArgs e)
        {
            if (gridTestTypes.CurrentCell == null) return;

            selectedTestType = gridTestTypes.CurrentRow.Cells["txtTestType"].Value.ToString();

            gridTestTypes.SelectionChanged -= gridTestTypes_SelectionChanged;
            gridTestTypes.CurrentRow.Selected = true;
            gridTestTypes.SelectionChanged += gridTestTypes_SelectionChanged;

            LoadTallies(gridTestTypes.CurrentRow.Cells["txtTestType"].Value.ToString());
        }

        private void LoadTallies(string TestType)
        {
            gridTallies.Rows.Clear();

            var tallies = monthlyReportEntity.TestTallies[selectedKey].Where(x => x.CategoryName == TestType).FirstOrDefault();

            categoryTallies = tallies.CategoryTallies;
            foreach(var item in tallies.CategoryTallies)
            {
                gridTallies.Rows.Add(item.SpecimenType, item.IP, item.ER, item.MAGS, item.OPD, item.Total);
            }

        }

        private void gridTallies_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex == -1) return;
            
            int x;
            if (gridTallies.Rows[e.RowIndex].Cells[e.ColumnIndex].Value == null)
            {
                gridTallies.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = 0;
                MessageBox.Show("Please Enter A Number!");
                return;
            }

            if(!Int32.TryParse(gridTallies.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString(), out x))
            {
                gridTallies.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = 0;
                MessageBox.Show("Please Enter A Number!");
                return;
            }

            string specimen = gridTallies.Rows[e.RowIndex].Cells["txtSpecimen"].Value.ToString();
            
            var specimenTally = categoryTallies.Where(y => y.SpecimenType == specimen).FirstOrDefault();

            specimenTally.IP = (e.ColumnIndex == 1) ? x : specimenTally.IP;
            specimenTally.ER = (e.ColumnIndex == 2) ? x : specimenTally.ER;
            specimenTally.MAGS = (e.ColumnIndex == 3) ? x : specimenTally.MAGS;
            specimenTally.OPD = (e.ColumnIndex == 4) ? x : specimenTally.OPD;

            gridTallies.CellValueChanged -= gridTallies_CellValueChanged;
            gridTallies.Rows[e.RowIndex].Cells["txtTotal"].Value = specimenTally.Total;
            gridTallies.CellValueChanged += gridTallies_CellValueChanged;
            RecalculateTestTypeSubTotals();

        }

        private void gridTallies_KeyDown(object sender, KeyEventArgs e)
        {

        }

        private void RecalculateTestTypeSubTotals()
        {
            for(int i = 0; i < gridTestTypes.Rows.Count; i++)
            {
                if(gridTestTypes.Rows[i].Cells["txtTestType"].Value.ToString() == selectedTestType)
                {
                    gridTestTypes.Rows[i].Cells["txtSubTotal"].Value = categoryTallies.Sum(x => x.Total);
                    gridTestTypes.Rows[i].Cells["IP"].Value = categoryTallies.Sum(x => x.IP);
                    gridTestTypes.Rows[i].Cells["ER"].Value = categoryTallies.Sum(x => x.ER);
                    gridTestTypes.Rows[i].Cells["MAGS"].Value = categoryTallies.Sum(x => x.MAGS);
                    gridTestTypes.Rows[i].Cells["OPD"].Value = categoryTallies.Sum(x => x.OPD);
                    break;
                }
            }

            for(int i = 0; i < gridWeeks.Rows.Count; i++)
            {
                if(gridWeeks.Rows[i].Cells["txtWeek"].Value.ToString() == selectedKey.ToString())
                {
                    var _tallies = monthlyReportEntity.TestTallies[selectedKey];
                    gridWeeks.Rows[i].Cells["txtWTotal"].Value = _tallies.Sum(x => x.CategoryTallyTotal);
                    gridWeeks.Rows[i].Cells["txtWIP"].Value = _tallies.Sum(x => x.CategoryTallies.Sum(y=>y.IP));
                    gridWeeks.Rows[i].Cells["txtWER"].Value = _tallies.Sum(x => x.CategoryTallies.Sum(y => y.ER));
                    gridWeeks.Rows[i].Cells["txtWMAGS"].Value = _tallies.Sum(x => x.CategoryTallies.Sum(y => y.MAGS));
                    gridWeeks.Rows[i].Cells["txtWOPD"].Value = _tallies.Sum(x => x.CategoryTallies.Sum(y => y.OPD));
                    break;
                }
            }
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            applicationSettings.AppForm.Show();
            this.Close();
        }

        private void txtWeek_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadTestTypes();
        }

        private void gridWeeks_SelectionChanged(object sender, EventArgs e)
        {
            if (gridWeeks.CurrentCell == null) return;
            selectedKey = (int)gridWeeks.CurrentRow.Cells["txtWeek"].Value;
            gridWeeks.CurrentRow.Selected = true;
            LoadTestTypes();
            gridTallies.CellValueChanged -= gridTallies_CellValueChanged;
            gridTallies.Rows.Clear();
            gridTallies.CellValueChanged += gridTallies_CellValueChanged;
        }
    }
}

using MonthlyReportEntrySoftware.Services;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MonthlyReportEntrySoftware
{
    public class ApplicationSettings
    {
        public string SavePath { get; set; }
        public Form AppForm { get; set; }

        public ApplicationSettings()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            ReportGeneratorService.BannerTitlecolor = (Color)ColorTranslator.FromHtml("#FFE598");
            ReportGeneratorService.ERColumnColor = (Color)ColorTranslator.FromHtml("#DEEAF6");
            ReportGeneratorService.IPColumnColor = (Color)ColorTranslator.FromHtml("#FFE598");
            ReportGeneratorService.MAGSColumnColor = (Color)ColorTranslator.FromHtml("#E2EFD9");
            ReportGeneratorService.OPColumnColor = (Color)ColorTranslator.FromHtml("#A9D08E");
            ReportGeneratorService.GrayColor = (Color)ColorTranslator.FromHtml("#DBDBDB");
        }
    }
}

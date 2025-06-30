using MonthlyReportEntrySoftware.Entities;
using MonthlyReportEntrySoftware.Entities.BaseEntities;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MonthlyReportEntrySoftware.Services
{
    public static class ReportGeneratorService
    {
        public static Color BannerTitlecolor { get; set; }
        public static Color ERColumnColor { get; set; }
        public static Color IPColumnColor { get; set; }
        public static Color MAGSColumnColor { get; set; }
        public static Color OPColumnColor { get; set; }
        public static Color GrayColor { get; set; }

        private static string[] colGroup = { "B", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y" };

        public static void GenerateMonthlyReport(MonthlyReportEntity monthlyReportEntity)
        {
            ExcelPackage package = new ExcelPackage(new FileInfo(monthlyReportEntity.FileName));

            var sheet = package.Workbook.Worksheets.Add("WEEKLY");
            var sheet2 = package.Workbook.Worksheets.Add("MONTHLY");

            #region Report Content

            InflateWeeklyReportRowContent(sheet, monthlyReportEntity);
            InflateMonthlyReportRowContent(sheet2, monthlyReportEntity);
            #endregion

            #region Monthly Report Styling

            sheet.Cells["E:Y"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            var bannerTitleRange = sheet.Cells[$"E1:W1,E31:W31,E57:W57,E95:W95"];
            bannerTitleRange.Merge = true;
            bannerTitleRange.Value = $"WEEKLY REPORT {DateTimeService.GetMonthFromCode(monthlyReportEntity.MonthCode)} {monthlyReportEntity.Year}";
            bannerTitleRange.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            bannerTitleRange.Style.Font.Size = 20;
            bannerTitleRange.Style.Fill.SetBackground(BannerTitlecolor);


            sheet.Cells["E3,J3,O3,T3,E33,J33,O33,T33,E59,J59,O59,T59,E97,J97,O97,T97"].Value = "IP";
            sheet.Cells["E33:E53,J33:J53,O33:O53,T33:T53,E97:E102,J97:J102,O97:O102,T97:T102"].Style.Fill.SetBackground(IPColumnColor);
            sheet.Cells["E90,J90,O90,T90"].Style.Fill.SetBackground(IPColumnColor);



            sheet.Cells["E3:E16,J3:J16,O3:O16,T3:T16"].Style.Fill.SetBackground(IPColumnColor);
            sheet.Cells["E20:E25,J20:J25,O20:O25,T20:T25"].Style.Fill.SetBackground(IPColumnColor);


            sheet.Cells["E61:E70,J61:J70,O61:O70,T61:T70"].Style.Fill.SetBackground(IPColumnColor);
            sheet.Cells["E73:E75,J73:J75,O73:O75,T73:T75"].Style.Fill.SetBackground(IPColumnColor);
            sheet.Cells["E77:E83,J77:J83,O77:O83,T77:T83"].Style.Fill.SetBackground(IPColumnColor);
            sheet.Cells["E86:E88,J86:J88,O86:O88,T86:T88"].Style.Fill.SetBackground(IPColumnColor);



            sheet.Cells["F3,K3,P3,U3,F33,K33,P33,U33,F59,K59,P59,U59,F97,K97,P97,U97"].Value = "ER";
            sheet.Cells["F33:F53,K33:K53,P33:P53,U33:U53,F97:F102,K97:K102,P97:P102,U97:U102"].Style.Fill.SetBackground(ERColumnColor);
            sheet.Cells["F90,K90,P90,U90"].Style.Fill.SetBackground(ERColumnColor);

            sheet.Cells["F3:F16,K3:K16,P3:P16,U3:U16"].Style.Fill.SetBackground(ERColumnColor);
            sheet.Cells["F20:F25,K20:K25,P20:P25,U20:U25"].Style.Fill.SetBackground(ERColumnColor);

            sheet.Cells["F61:F70,K61:K70,P61:P70,U61:U70"].Style.Fill.SetBackground(ERColumnColor);
            sheet.Cells["F73:F75,K73:K75,P73:P75,U73:U75"].Style.Fill.SetBackground(ERColumnColor);
            sheet.Cells["F77:F83,K77:K83,P77:P83,U77:U83"].Style.Fill.SetBackground(ERColumnColor);
            sheet.Cells["F86:F88,K86:K88,P86:P88,U86:U88"].Style.Fill.SetBackground(ERColumnColor);

            // === MAGS Column ===
            sheet.Cells["G3,L3,Q3,V3,G33,L33,Q33,V33,G59,L59,Q59,V59,G97,L97,Q97,V97"].Value = "MAGS";
            sheet.Cells["G33:G53,L33:L53,Q33:Q53,V33:V53,G97:G102,L97:L102,Q97:Q102,V97:V102"].Style.Fill.SetBackground(MAGSColumnColor);
            sheet.Cells["G90,L90,Q90,V90"].Style.Fill.SetBackground(MAGSColumnColor);

            sheet.Cells["G3:G16,L3:L16,Q3:Q16,V3:V16"].Style.Fill.SetBackground(MAGSColumnColor);
            sheet.Cells["G20:G25,L20:L25,Q20:Q25,V20:V25"].Style.Fill.SetBackground(MAGSColumnColor);

            sheet.Cells["G61:G70,L61:L70,Q61:Q70,V61:V70"].Style.Fill.SetBackground(MAGSColumnColor);
            sheet.Cells["G73:G75,L73:L75,Q73:Q75,V73:V75"].Style.Fill.SetBackground(MAGSColumnColor);
            sheet.Cells["G77:G83,L77:L83,Q77:Q83,V77:V83"].Style.Fill.SetBackground(MAGSColumnColor);
            sheet.Cells["G86:G88,L86:L88,Q86:Q88,V86:V88"].Style.Fill.SetBackground(MAGSColumnColor);

            // === OP Column ===
            sheet.Cells["H3,M3,R3,W3,H33,M33,R33,W33,H59,M59,R59,W59,H97,M97,R97,W97"].Value = "OP";
            sheet.Cells["H33:H53,M33:M53,R33:R53,W33:W53,H97:H102,M97:M102,R97:R102,W97:W102"].Style.Fill.SetBackground(OPColumnColor);
            sheet.Cells["H90,M90,R90,W90"].Style.Fill.SetBackground(OPColumnColor);

            sheet.Cells["H3:H16,M3:M16,R3:R16,W3:W16"].Style.Fill.SetBackground(OPColumnColor);
            sheet.Cells["H20:H25,M20:M25,R20:R25,W20:W25"].Style.Fill.SetBackground(OPColumnColor);

            sheet.Cells["H61:H70,M61:M70,R61:R70,W61:W70"].Style.Fill.SetBackground(OPColumnColor);
            sheet.Cells["H73:H75,M73:M75,R73:R75,W73:W75"].Style.Fill.SetBackground(OPColumnColor);
            sheet.Cells["H77:H83,M77:M83,R77:R83,W77:W83"].Style.Fill.SetBackground(OPColumnColor);
            sheet.Cells["H86:H88,M86:M88,R86:R88,W86:W88"].Style.Fill.SetBackground(OPColumnColor);


            sheet.Cells["B2:Y28,B32:Y54,B58:Y92,B96:Y103"].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
            sheet.Cells["B2:D3"].Merge = true;
            sheet.Cells["B2:D3"].Value = "BACTERIOLOGY";
            sheet.Cells["B2:D3"].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
            sheet.Cells["B2:D3"].Style.Font.Bold = true;
            sheet.Cells["B2:D3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Cells["B2:D3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            sheet.Cells["B32:D33"].Merge = true;
            sheet.Cells["B32:D33"].Value = "GRAM STAINING";
            sheet.Cells["B32:D33"].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
            sheet.Cells["B32:D33"].Style.Font.Bold = true;
            sheet.Cells["B32:D33"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Cells["B32:D33"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            sheet.Cells["B58:D59"].Merge = true;
            sheet.Cells["B58:D59"].Value = "CULTURE AND SENSI";
            sheet.Cells["B58:D59"].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
            sheet.Cells["B58:D59"].Style.Font.Bold = true;
            sheet.Cells["B58:D59"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Cells["B58:D59"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            sheet.Cells["B96:D97"].Merge = true;
            sheet.Cells["B96:D97"].Value = "OTHER TESTS";
            sheet.Cells["B96:D97"].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
            sheet.Cells["B96:D97"].Style.Font.Bold = true;
            sheet.Cells["B96:D97"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Cells["B96:D97"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            var tempRange = sheet.Cells["E2:H2,E32:H32,E58:H58,E96:H96"];
            tempRange.Value = "1ST WK";
            tempRange.Merge = true;
            tempRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            tempRange.Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);

            tempRange = sheet.Cells["J2:M2,J32:M32,J58:M58,J96:M96"];
            tempRange.Value = "2ND WK";
            tempRange.Merge = true;
            tempRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            tempRange.Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);

            tempRange = sheet.Cells["O2:R2,O32:R32,O58:R58,O96:R96"];
            tempRange.Value = "3RD WK";
            tempRange.Merge = true;
            tempRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            tempRange.Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);

            tempRange = sheet.Cells["T2:W2,T32:W32,T58:W58,T96:W96"];
            tempRange.Value = "4TH WK";
            tempRange.Merge = true;
            tempRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            tempRange.Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);

            sheet.Cells["Y2,Y32,Y58,Y97"].Value = "TOTAL";


            sheet.Cells[$"B26:Y26"].Style.Fill.SetBackground(Color.Gold);
            #endregion

         

            Stream stream = File.Create(monthlyReportEntity.FullPath);
            package.SaveAs(stream);
            stream.Close();
        }
        public static void LoadFromExcelFile(MonthlyReportEntity monthlyReportEntity)
        {
            using (ExcelPackage package = new ExcelPackage(new FileInfo(monthlyReportEntity.FullPath)))
            {
                var sheet = package.Workbook.Worksheets[0];

                #region AFB

                try
                {

                    for (int i = 5; i <= 25; i++)
                    {

                        if (i == 17 || i == 18 || i == 19) continue;
                        string currentSpecimen = sheet.Cells[$"C{i}"].Text;
                        int _index = ((currentSpecimen.IndexOf(".") < 0) ? 0 : currentSpecimen.IndexOf(" "));
                        currentSpecimen = currentSpecimen.Substring(_index).Trim();
                        
                        monthlyReportEntity.TestTallies[1][0].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().IP = ZeroIfEmpty(sheet.Cells[$"E{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[2][0].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().IP = ZeroIfEmpty(sheet.Cells[$"J{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[3][0].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().IP = ZeroIfEmpty(sheet.Cells[$"O{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[4][0].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().IP = ZeroIfEmpty(sheet.Cells[$"T{i}"].Value?.ToString());

                        monthlyReportEntity.TestTallies[1][0].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().ER = ZeroIfEmpty(sheet.Cells[$"F{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[2][0].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().ER = ZeroIfEmpty(sheet.Cells[$"K{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[3][0].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().ER = ZeroIfEmpty(sheet.Cells[$"P{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[4][0].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().ER = ZeroIfEmpty(sheet.Cells[$"U{i}"].Value?.ToString());

                        monthlyReportEntity.TestTallies[1][0].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().MAGS = ZeroIfEmpty(sheet.Cells[$"G{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[2][0].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().MAGS = ZeroIfEmpty(sheet.Cells[$"L{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[3][0].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().MAGS = ZeroIfEmpty(sheet.Cells[$"Q{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[4][0].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().MAGS = ZeroIfEmpty(sheet.Cells[$"V{i}"].Value?.ToString());

                        monthlyReportEntity.TestTallies[1][0].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().OPD = ZeroIfEmpty(sheet.Cells[$"H{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[2][0].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().OPD = ZeroIfEmpty(sheet.Cells[$"M{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[3][0].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().OPD = ZeroIfEmpty(sheet.Cells[$"R{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[4][0].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().OPD = ZeroIfEmpty(sheet.Cells[$"W{i}"].Value?.ToString());

                    }

                    #endregion

                    #region GS

                    for (int i = 34; i <= 53; i++)
                    {

                        string currentSpecimen = sheet.Cells[$"C{i}"].Text;
                        int _index = ((currentSpecimen.IndexOf(".") < 0) ? 0 : currentSpecimen.IndexOf(" "));
                        currentSpecimen = currentSpecimen.Substring(_index).Trim();

                        monthlyReportEntity.TestTallies[1][1].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().IP = ZeroIfEmpty(sheet.Cells[$"E{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[2][1].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().IP = ZeroIfEmpty(sheet.Cells[$"J{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[3][1].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().IP = ZeroIfEmpty(sheet.Cells[$"O{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[4][1].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().IP = ZeroIfEmpty(sheet.Cells[$"T{i}"].Value?.ToString());

                        monthlyReportEntity.TestTallies[1][1].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().ER = ZeroIfEmpty(sheet.Cells[$"F{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[2][1].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().ER = ZeroIfEmpty(sheet.Cells[$"K{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[3][1].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().ER = ZeroIfEmpty(sheet.Cells[$"P{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[4][1].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().ER = ZeroIfEmpty(sheet.Cells[$"U{i}"].Value?.ToString());

                        monthlyReportEntity.TestTallies[1][1].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().MAGS = ZeroIfEmpty(sheet.Cells[$"G{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[2][1].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().MAGS = ZeroIfEmpty(sheet.Cells[$"L{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[3][1].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().MAGS = ZeroIfEmpty(sheet.Cells[$"Q{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[4][1].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().MAGS = ZeroIfEmpty(sheet.Cells[$"V{i}"].Value?.ToString());

                        monthlyReportEntity.TestTallies[1][1].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().OPD = ZeroIfEmpty(sheet.Cells[$"H{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[2][1].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().OPD = ZeroIfEmpty(sheet.Cells[$"M{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[3][1].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().OPD = ZeroIfEmpty(sheet.Cells[$"R{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[4][1].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().OPD = ZeroIfEmpty(sheet.Cells[$"W{i}"].Value?.ToString());

                    }

                    #endregion

                    #region CS

                    for (int i = 61; i <= 90; i++)
                    {
                        if (i == 71 || i == 72 || i == 76 || i == 84 || i == 85 || i == 89) continue;
                        string currentSpecimen = sheet.Cells[$"C{i}"].Text;
                        int _index = ((currentSpecimen.IndexOf(".") < 0) ? 0 : currentSpecimen.IndexOf(" "));
                        currentSpecimen = currentSpecimen.Substring(_index).Trim();

                        monthlyReportEntity.TestTallies[1][2].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().IP = ZeroIfEmpty(sheet.Cells[$"E{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[2][2].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().IP = ZeroIfEmpty(sheet.Cells[$"J{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[3][2].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().IP = ZeroIfEmpty(sheet.Cells[$"O{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[4][2].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().IP = ZeroIfEmpty(sheet.Cells[$"T{i}"].Value?.ToString());

                        monthlyReportEntity.TestTallies[1][2].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().ER = ZeroIfEmpty(sheet.Cells[$"F{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[2][2].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().ER = ZeroIfEmpty(sheet.Cells[$"K{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[3][2].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().ER = ZeroIfEmpty(sheet.Cells[$"P{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[4][2].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().ER = ZeroIfEmpty(sheet.Cells[$"U{i}"].Value?.ToString());

                        monthlyReportEntity.TestTallies[1][2].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().MAGS = ZeroIfEmpty(sheet.Cells[$"G{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[2][2].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().MAGS = ZeroIfEmpty(sheet.Cells[$"L{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[3][2].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().MAGS = ZeroIfEmpty(sheet.Cells[$"Q{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[4][2].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().MAGS = ZeroIfEmpty(sheet.Cells[$"V{i}"].Value?.ToString());

                        monthlyReportEntity.TestTallies[1][2].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().OPD = ZeroIfEmpty(sheet.Cells[$"H{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[2][2].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().OPD = ZeroIfEmpty(sheet.Cells[$"M{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[3][2].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().OPD = ZeroIfEmpty(sheet.Cells[$"R{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[4][2].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim() && !x.IsHeader).FirstOrDefault().OPD = ZeroIfEmpty(sheet.Cells[$"W{i}"].Value?.ToString());

                    }

                    #endregion

                    #region Other Tests

                    for (int i = 98; i <= 102; i++)
                    {
                        string currentSpecimen = sheet.Cells[$"C{i}"].Text;

                        monthlyReportEntity.TestTallies[1][i - 95].CategoryTallies.FirstOrDefault().IP = ZeroIfEmpty(sheet.Cells[$"E{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[2][i - 95].CategoryTallies.FirstOrDefault().IP = ZeroIfEmpty(sheet.Cells[$"J{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[3][i - 95].CategoryTallies.FirstOrDefault().IP = ZeroIfEmpty(sheet.Cells[$"O{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[4][i - 95].CategoryTallies.FirstOrDefault().IP = ZeroIfEmpty(sheet.Cells[$"T{i}"].Value?.ToString());

                        monthlyReportEntity.TestTallies[1][i - 95].CategoryTallies.FirstOrDefault().ER = ZeroIfEmpty(sheet.Cells[$"F{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[2][i - 95].CategoryTallies.FirstOrDefault().ER = ZeroIfEmpty(sheet.Cells[$"K{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[3][i - 95].CategoryTallies.FirstOrDefault().ER = ZeroIfEmpty(sheet.Cells[$"P{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[4][i - 95].CategoryTallies.FirstOrDefault().ER = ZeroIfEmpty(sheet.Cells[$"U{i}"].Value?.ToString());

                        monthlyReportEntity.TestTallies[1][i - 95].CategoryTallies.FirstOrDefault().MAGS = ZeroIfEmpty(sheet.Cells[$"G{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[2][i - 95].CategoryTallies.FirstOrDefault().MAGS = ZeroIfEmpty(sheet.Cells[$"L{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[3][i - 95].CategoryTallies.FirstOrDefault().MAGS = ZeroIfEmpty(sheet.Cells[$"Q{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[4][i - 95].CategoryTallies.FirstOrDefault().MAGS = ZeroIfEmpty(sheet.Cells[$"V{i}"].Value?.ToString());

                        monthlyReportEntity.TestTallies[1][i - 95].CategoryTallies.FirstOrDefault().OPD = ZeroIfEmpty(sheet.Cells[$"H{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[2][i - 95].CategoryTallies.FirstOrDefault().OPD = ZeroIfEmpty(sheet.Cells[$"M{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[3][i - 95].CategoryTallies.FirstOrDefault().OPD = ZeroIfEmpty(sheet.Cells[$"R{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[4][i - 95].CategoryTallies.FirstOrDefault().OPD = ZeroIfEmpty(sheet.Cells[$"W{i}"].Value?.ToString());

                    }

                    #endregion

                }
                catch (Exception ex)
                {
                    throw new Exception("Malformed Excel File!");
                }



            }
        }
        private static void InflateWeeklyReportRowContent(ExcelWorksheet sheet, MonthlyReportEntity monthlyReportEntity)
        {   

            int row = 5;

            sheet.Cells[$"B4:D4"].Merge = true;
            sheet.Cells[$"B4:D4"].Value = "AFB";
            sheet.Cells[$"B4:D4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Cells[$"B4:D4"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            sheet.Cells[$"B1:Y103"].Style.Font.Bold = true;
            sheet.Cells[$"B1:Y103"].Style.Font.Size = 10;

            sheet.Cells["I5:I17,I20:I26,I34:I53,I61:I92"].Style.Font.Color.SetColor(Color.Red);
            sheet.Cells["N5:N17,N20:N26,N34:N53,N61:N92"].Style.Font.Color.SetColor(Color.Red);
            sheet.Cells["S5:S17,S20:S26,S34:S53,S61:S92"].Style.Font.Color.SetColor(Color.Red);
            sheet.Cells["X5:X17,X20:X26,X34:X53,X61:X92"].Style.Font.Color.SetColor(Color.Red);
            sheet.Cells["Y5:Y17,Y20:Y26,Y34:Y53,Y61:Y92"].Style.Font.Color.SetColor(Color.Blue);

            sheet.Cells["B2:Y28,B32:Y54,B58:Y92,B96:Y103"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            sheet.Cells["B2:Y28,B32:Y54,B58:Y92,B96:Y103"].Style.Border.Top.Color.SetColor(Color.Black);

            sheet.Cells["B2:Y28,B32:Y54,B58:Y92,B96:Y103"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            sheet.Cells["B2:Y28,B32:Y54,B58:Y92,B96:Y103"].Style.Border.Bottom.Color.SetColor(Color.Black);

            sheet.Cells["B2:Y28,B32:Y54,B58:Y92,B96:Y103"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            sheet.Cells["B2:Y28,B32:Y54,B58:Y92,B96:Y103"].Style.Border.Left.Color.SetColor(Color.Black);

            sheet.Cells["B2:Y28,B32:Y54,B58:Y92,B96:Y103"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            sheet.Cells["B2:Y28,B32:Y54,B58:Y92,B96:Y103"].Style.Border.Right.Color.SetColor(Color.Black);


            var tallies = monthlyReportEntity.TestTallies[1][0].CategoryTallies.ToList();

            #region AFB Category
            foreach (var tally in tallies)
            {
                var specimen = tally.SpecimenType;

                string counterLabel = (row - 4).ToString() + ". ";
                sheet.Cells[$"B{row}:D{row}"].Merge = true;
            

                if (row - 4 == 13)
                {
                    sheet.Cells[$"B{row}:Y{row}"].Style.Fill.SetBackground(Color.Gold);


                    sheet.Cells[$"B{row}:D{row}"].Merge = true;
                    sheet.Cells[$"B{row}:D{row}"].Style.Font.Bold = false;
                    sheet.Cells[$"B{row}:D{row}"].Style.Font.Italic = true;
                    sheet.Cells[$"B{row}:D{row}"].Value = "SUB-TOTAL: ";

                    sheet.Cells[$"E{row}"].Formula = "=SUM(E5:E16)";
                    sheet.Cells[$"F{row}"].Formula = "=SUM(F5:F16)";
                    sheet.Cells[$"G{row}"].Formula = "=SUM(G5:G16)";
                    sheet.Cells[$"H{row}"].Formula = "=SUM(H5:H16)";


                    sheet.Cells[$"J{row}"].Formula = "=SUM(J5:J16)";
                    sheet.Cells[$"K{row}"].Formula = "=SUM(K5:K16)";
                    sheet.Cells[$"L{row}"].Formula = "=SUM(L5:L16)";
                    sheet.Cells[$"M{row}"].Formula = "=SUM(M5:M16)";


                    sheet.Cells[$"O{row}"].Formula = "=SUM(O5:O16)";
                    sheet.Cells[$"P{row}"].Formula = "=SUM(P5:P16)";
                    sheet.Cells[$"Q{row}"].Formula = "=SUM(Q5:Q16)";
                    sheet.Cells[$"R{row}"].Formula = "=SUM(R5:R16)";


                    sheet.Cells[$"T{row}"].Formula = "=SUM(T5:T16)";
                    sheet.Cells[$"U{row}"].Formula = "=SUM(U5:U16)";
                    sheet.Cells[$"V{row}"].Formula = "=SUM(V5:V16)";
                    sheet.Cells[$"W{row}"].Formula = "=SUM(W5:W16)";

                    sheet.Cells[$"I{row}"].Formula = "=SUM(I5:I16)";
                    sheet.Cells[$"N{row}"].Formula = "=SUM(N5:N16)";
                    sheet.Cells[$"S{row}"].Formula = "=SUM(S5:S16)";
                    sheet.Cells[$"X{row}"].Formula = "=SUM(X5:X16)";


                    sheet.Cells[$"Y{row}"].Formula = $"=SUM(Y5:Y16)";


                    row++;
                    sheet.Cells[$"B{row}:D{row}"].Merge = true;
                    row++;
                }


                if (row >= 19)
                {
                    counterLabel = string.Empty;
                    sheet.Cells[$"B{row}:D{row}"].Style.Fill.SetBackground(GrayColor);
                }

                if (tally.IsHeader)
                {
                    counterLabel = string.Empty;
                    sheet.Cells[$"B{row}:D{row}"].Merge = true;
                    sheet.Cells[$"B{row}:D{row}"].Value = $"{specimen}";
                    sheet.Cells[$"B{row}:Y{row}"].Style.Fill.SetBackground(Color.Yellow);
                    sheet.Cells[$"B{row}:Y{row}"].Style.Font.Color.SetColor(Color.Red);
                    row++;
                    continue;
                }
                else
                {
                    sheet.Cells[$"B{row}:D{row}"].Value = $"{counterLabel} {specimen}";
                }





                sheet.Cells[$"E{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[1][0].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().IP);
                sheet.Cells[$"J{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[2][0].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().IP);
                sheet.Cells[$"O{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[3][0].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().IP);
                sheet.Cells[$"T{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[4][0].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().IP);

                sheet.Cells[$"F{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[1][0].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().ER);
                sheet.Cells[$"K{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[2][0].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().ER);
                sheet.Cells[$"P{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[3][0].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().ER);
                sheet.Cells[$"U{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[4][0].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().ER);

                sheet.Cells[$"G{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[1][0].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().MAGS);
                sheet.Cells[$"L{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[2][0].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().MAGS);
                sheet.Cells[$"Q{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[3][0].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().MAGS);
                sheet.Cells[$"V{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[4][0].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().MAGS);

                sheet.Cells[$"H{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[1][0].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().OPD);
                sheet.Cells[$"M{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[2][0].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().OPD);
                sheet.Cells[$"R{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[3][0].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().OPD);
                sheet.Cells[$"W{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[4][0].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().OPD);

                sheet.Cells[$"I{row}"].Formula = $"=SUM(E{row},F{row},G{row},H{row})";
                sheet.Cells[$"N{row}"].Formula = $"=SUM(J{row},K{row},L{row},M{row})";
                sheet.Cells[$"S{row}"].Formula = $"=SUM(O{row},P{row},Q{row},R{row})";
                sheet.Cells[$"X{row}"].Formula = $"=SUM(T{row},U{row},V{row},W{row})";

                sheet.Cells[$"Y{row}"].Formula = $"=SUM(I{row},N{row},S{row},X{row})";

                row++;

            }

            sheet.Cells[$"B{row}:D{row}"].Merge = true;
            sheet.Cells[$"B{row}:D{row}"].Style.Font.Bold = false;
            sheet.Cells[$"B{row}:D{row}"].Style.Font.Italic = true;
            sheet.Cells[$"B{row}:D{row}"].Value = "TOTAL NO. OF OTHERS: ";

            sheet.Cells[$"E{row}"].Formula = "=SUM(E20:E25)";
            sheet.Cells[$"F{row}"].Formula = "=SUM(F20:F25)";
            sheet.Cells[$"G{row}"].Formula = "=SUM(G20:G25)";
            sheet.Cells[$"H{row}"].Formula = "=SUM(H20:H25)";


            sheet.Cells[$"J{row}"].Formula = "=SUM(J20:J25)";
            sheet.Cells[$"K{row}"].Formula = "=SUM(K20:K25)";
            sheet.Cells[$"L{row}"].Formula = "=SUM(L20:L25)";
            sheet.Cells[$"M{row}"].Formula = "=SUM(M20:M25)";


            sheet.Cells[$"O{row}"].Formula = "=SUM(O20:O25)";
            sheet.Cells[$"P{row}"].Formula = "=SUM(P20:P25)";
            sheet.Cells[$"Q{row}"].Formula = "=SUM(Q20:Q25)";
            sheet.Cells[$"R{row}"].Formula = "=SUM(R20:R25)";


            sheet.Cells[$"T{row}"].Formula = "=SUM(T20:T25)";
            sheet.Cells[$"U{row}"].Formula = "=SUM(U20:U25)";
            sheet.Cells[$"V{row}"].Formula = "=SUM(V20:V25)";
            sheet.Cells[$"W{row}"].Formula = "=SUM(W20:W25)";

            sheet.Cells[$"I{row}"].Formula = "=SUM(I20:I25)";
            sheet.Cells[$"N{row}"].Formula = "=SUM(N20:N25)";
            sheet.Cells[$"S{row}"].Formula = "=SUM(S20:S25)";
            sheet.Cells[$"X{row}"].Formula = "=SUM(X20:X25)";

            sheet.Cells[$"Y{row}"].Formula = "=SUM(I27,N27,S27,X27)";

            sheet.Cells[$"I{row}"].Style.Font.Color.SetColor(Color.Red);
            sheet.Cells[$"N{row}"].Style.Font.Color.SetColor(Color.Red);
            sheet.Cells[$"S{row}"].Style.Font.Color.SetColor(Color.Red);
            sheet.Cells[$"X{row}"].Style.Font.Color.SetColor(Color.Red);
            sheet.Cells[$"Y{row}"].Style.Font.Color.SetColor(Color.Blue);

            sheet.Cells["B27:D27"].Merge = true;

            sheet.Cells["B28:D28"].Merge = true;
            sheet.Cells["B28:D28"].Value = "TOTAL NO. OF AFB RECEIVED:";

            sheet.Cells[$"E28"].Formula = "=SUM(E17,E26)";
            sheet.Cells[$"F28"].Formula = "=SUM(F17,F26)";
            sheet.Cells[$"G28"].Formula = "=SUM(G17,G26)";
            sheet.Cells[$"H28"].Formula = "=SUM(H17,H26)";


            sheet.Cells[$"J28"].Formula = "=SUM(J17,J26)";
            sheet.Cells[$"K28"].Formula = "=SUM(K17,K26)";
            sheet.Cells[$"L28"].Formula = "=SUM(L17,L26)";
            sheet.Cells[$"M28"].Formula = "=SUM(M17,M26)";


            sheet.Cells[$"O28"].Formula = "=SUM(O17,O26)";
            sheet.Cells[$"P28"].Formula = "=SUM(P17,P26)";
            sheet.Cells[$"Q28"].Formula = "=SUM(Q17,Q26)";
            sheet.Cells[$"R28"].Formula = "=SUM(R17,R26)";


            sheet.Cells[$"T28"].Formula = "=SUM(T17,T26)";
            sheet.Cells[$"U28"].Formula = "=SUM(U17,U26)";
            sheet.Cells[$"V28"].Formula = "=SUM(V17,V26)";
            sheet.Cells[$"W28"].Formula = "=SUM(W17,W26)";


            sheet.Cells["I28"].Formula = "SUM(I17,I26)";
            sheet.Cells["N28"].Formula = "SUM(N17,N26)";
            sheet.Cells["S28"].Formula = "SUM(S17,S26)";
            sheet.Cells["X28"].Formula = "SUM(X17,X26)";

            sheet.Cells["Y28"].Formula = "SUM(Y17,Y26)";
            sheet.Cells["Y28"].Style.Font.Color.SetColor(Color.Purple);
            sheet.Cells["B28:Y28"].Style.Border.BorderAround(ExcelBorderStyle.Thick, Color.Black);

            #endregion

            #region GS Category
            row = 34;

            tallies = monthlyReportEntity.TestTallies[1][1].CategoryTallies.ToList();

            foreach (var tally in tallies)
            {
                var specimen = tally.SpecimenType;
                string counterLabel = (row - 33).ToString() + ". ";
                sheet.Cells[$"B{row}:D{row}"].Merge = true;

                sheet.Cells[$"B{row}:D{row}"].Value = $"{counterLabel} {specimen}";

                sheet.Cells[$"E{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[1][1].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().IP);
                sheet.Cells[$"J{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[2][1].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().IP);
                sheet.Cells[$"O{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[3][1].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().IP);
                sheet.Cells[$"T{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[4][1].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().IP);

                sheet.Cells[$"F{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[1][1].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().ER);
                sheet.Cells[$"K{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[2][1].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().ER);
                sheet.Cells[$"P{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[3][1].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().ER);
                sheet.Cells[$"U{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[4][1].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().ER);

                sheet.Cells[$"G{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[1][1].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().MAGS);
                sheet.Cells[$"L{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[2][1].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().MAGS);
                sheet.Cells[$"Q{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[3][1].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().MAGS);
                sheet.Cells[$"V{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[4][1].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().MAGS);

                sheet.Cells[$"H{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[1][1].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().OPD);
                sheet.Cells[$"M{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[2][1].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().OPD);
                sheet.Cells[$"R{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[3][1].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().OPD);
                sheet.Cells[$"W{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[4][1].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().OPD);

                sheet.Cells[$"I{row}"].Formula = $"=SUM(E{row},F{row},G{row},H{row})";
                sheet.Cells[$"N{row}"].Formula = $"=SUM(J{row},K{row},L{row},M{row})";
                sheet.Cells[$"S{row}"].Formula = $"=SUM(O{row},P{row},Q{row},R{row})";
                sheet.Cells[$"X{row}"].Formula = $"=SUM(T{row},U{row},V{row},W{row})";

                sheet.Cells[$"Y{row}"].Formula = $"=SUM(I{row},N{row},S{row},X{row})";

                row++;
            }

            sheet.Cells[$"B46:D46"].Merge = true;
            row++;
            sheet.Cells[$"B54:D54"].Merge = true;
            sheet.Cells[$"B54:D54"].Value = "TOTAL NO OF. SPECIMEN RECEIVED:";
            sheet.Cells["B54:Y54"].Style.Border.BorderAround(ExcelBorderStyle.Thick, Color.Black);

            sheet.Cells[$"E54"].Formula = "=SUM(E34:E53)";
            sheet.Cells[$"F54"].Formula = "=SUM(F34:F53)";
            sheet.Cells[$"G54"].Formula = "=SUM(G34:G53)";
            sheet.Cells[$"H54"].Formula = "=SUM(H34:H53)";


            sheet.Cells[$"J54"].Formula = "=SUM(J34:J53)";
            sheet.Cells[$"K54"].Formula = "=SUM(K34:K53)";
            sheet.Cells[$"L54"].Formula = "=SUM(L34:L53)";
            sheet.Cells[$"M54"].Formula = "=SUM(M34:M53)";


            sheet.Cells[$"O54"].Formula = "=SUM(O34:O53)";
            sheet.Cells[$"P54"].Formula = "=SUM(P34:P53)";
            sheet.Cells[$"Q54"].Formula = "=SUM(Q34:Q53)";
            sheet.Cells[$"R54"].Formula = "=SUM(R34:R53)";


            sheet.Cells[$"T54"].Formula = "=SUM(T34:T53)";
            sheet.Cells[$"U54"].Formula = "=SUM(U34:U53)";
            sheet.Cells[$"V54"].Formula = "=SUM(V34:V53)";
            sheet.Cells[$"W54"].Formula = "=SUM(W34:W53)";

            sheet.Cells["I54"].Formula = "=SUM(I34:I53)";
            sheet.Cells["N54"].Formula = "=SUM(N34:N53)";
            sheet.Cells["S54"].Formula = "=SUM(S34:S53)";
            sheet.Cells["X54"].Formula = "=SUM(X34:X53)";
            sheet.Cells["Y54"].Formula = "=SUM(Y34:Y53)";

            sheet.Cells["I54"].Style.Font.Color.SetColor(Color.Red);
            sheet.Cells["N54"].Style.Font.Color.SetColor(Color.Red);
            sheet.Cells["S54"].Style.Font.Color.SetColor(Color.Red);
            sheet.Cells["X54"].Style.Font.Color.SetColor(Color.Red);
            sheet.Cells["Y54"].Style.Font.Color.SetColor(Color.Purple);

            #endregion

            #region C/S Category
            row = 60;

            tallies = monthlyReportEntity.TestTallies[1][2].CategoryTallies.ToList();
            int ctrLabel = 1, incrementFactor = 0;
            foreach (var tally in tallies)
            {
                var specimen = tally.SpecimenType;
                string specimenLabel = $"{ctrLabel}. {specimen}";
                sheet.Cells[$"B{row}:D{row}"].Merge = true;

                

                if (row == 71)
                {
                    sheet.Cells[$"B{row}:D{row}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    sheet.Cells[$"B{row}:Y{row}"].Style.Fill.SetBackground(Color.Gold);


                    sheet.Cells[$"B{row}:D{row}"].Merge = true;
                    sheet.Cells[$"B{row}:D{row}"].Style.Font.Bold = false;
                    sheet.Cells[$"B{row}:D{row}"].Style.Font.Italic = true;
                    sheet.Cells[$"B{row}:D{row}"].Value = "TOTAL TRANSUDATES: ";

                    sheet.Cells[$"E{row}"].Formula = "=SUM(E61:E70)";
                    sheet.Cells[$"F{row}"].Formula = "=SUM(F61:F70)";
                    sheet.Cells[$"G{row}"].Formula = "=SUM(G61:G70)";
                    sheet.Cells[$"H{row}"].Formula = "=SUM(H61:H70)";


                    sheet.Cells[$"J{row}"].Formula = "=SUM(J61:J70)";
                    sheet.Cells[$"K{row}"].Formula = "=SUM(K61:K70)";
                    sheet.Cells[$"L{row}"].Formula = "=SUM(L61:L70)";
                    sheet.Cells[$"M{row}"].Formula = "=SUM(M61:M70)";


                    sheet.Cells[$"O{row}"].Formula = "=SUM(O61:O70)";
                    sheet.Cells[$"P{row}"].Formula = "=SUM(P61:P70)";
                    sheet.Cells[$"Q{row}"].Formula = "=SUM(Q61:Q70)";
                    sheet.Cells[$"R{row}"].Formula = "=SUM(R61:R70)";


                    sheet.Cells[$"T{row}"].Formula = "=SUM(T61:T70)";
                    sheet.Cells[$"U{row}"].Formula = "=SUM(U61:U70)";
                    sheet.Cells[$"V{row}"].Formula = "=SUM(V61:V70)";
                    sheet.Cells[$"W{row}"].Formula = "=SUM(W61:W70)";

                    sheet.Cells[$"I{row}"].Formula = "=SUM(I61:I70)";
                    sheet.Cells[$"N{row}"].Formula = "=SUM(N61:N70)";
                    sheet.Cells[$"S{row}"].Formula = "=SUM(S61:S70)";
                    sheet.Cells[$"X{row}"].Formula = "=SUM(X61:X70)";


                    sheet.Cells[$"Y{row}"].Formula = $"=SUM(Y61:Y70)";

                    row++;
                    sheet.Cells[$"B{row}:D{row}"].Merge = true;
                    row++;
                    ctrLabel++;
                    incrementFactor = 1;
                    specimenLabel = $"{ctrLabel}. {specimen}";
                }



                if (row == 84)
                {
                    sheet.Cells[$"B{row}:D{row}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    sheet.Cells[$"B{row}:Y{row}"].Style.Fill.SetBackground(Color.Gold);


                    sheet.Cells[$"B{row}:D{row}"].Merge = true;
                    sheet.Cells[$"B{row}:D{row}"].Style.Font.Bold = false;
                    sheet.Cells[$"B{row}:D{row}"].Style.Font.Italic = true;
                    sheet.Cells[$"B{row}:D{row}"].Value = "TOTAL DISCHARGES: ";

                    sheet.Cells[$"E{row}"].Formula = "=SUM(E77:E83)";
                    sheet.Cells[$"F{row}"].Formula = "=SUM(F77:F83)";
                    sheet.Cells[$"G{row}"].Formula = "=SUM(G77:G83)";
                    sheet.Cells[$"H{row}"].Formula = "=SUM(H77:H83)";


                    sheet.Cells[$"J{row}"].Formula = "=SUM(J77:J83)";
                    sheet.Cells[$"K{row}"].Formula = "=SUM(K77:K83)";
                    sheet.Cells[$"L{row}"].Formula = "=SUM(L77:L83)";
                    sheet.Cells[$"M{row}"].Formula = "=SUM(M77:M83)";


                    sheet.Cells[$"O{row}"].Formula = "=SUM(O77:O83)";
                    sheet.Cells[$"P{row}"].Formula = "=SUM(P77:P83)";
                    sheet.Cells[$"Q{row}"].Formula = "=SUM(Q77:Q83)";
                    sheet.Cells[$"R{row}"].Formula = "=SUM(R77:R83)";


                    sheet.Cells[$"T{row}"].Formula = "=SUM(T77:T83)";
                    sheet.Cells[$"U{row}"].Formula = "=SUM(U77:U83)";
                    sheet.Cells[$"V{row}"].Formula = "=SUM(V77:V83)";
                    sheet.Cells[$"W{row}"].Formula = "=SUM(W77:W83)";

                    sheet.Cells[$"I{row}"].Formula = "=SUM(I77:I83)";
                    sheet.Cells[$"N{row}"].Formula = "=SUM(N77:N83)";
                    sheet.Cells[$"S{row}"].Formula = "=SUM(S77:S83)";
                    sheet.Cells[$"X{row}"].Formula = "=SUM(X77:X83)";


                    sheet.Cells[$"Y{row}"].Formula = $"=SUM(Y77:Y83)";

                    row++;
                    ctrLabel++;
                }


                if (row == 89)
                {
                    sheet.Cells[$"B{row}:D{row}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    sheet.Cells[$"B{row}:Y{row}"].Style.Fill.SetBackground(Color.Gold);


                    sheet.Cells[$"B{row}:D{row}"].Merge = true;
                    sheet.Cells[$"B{row}:D{row}"].Style.Font.Bold = false;
                    sheet.Cells[$"B{row}:D{row}"].Style.Font.Italic = true;
                    sheet.Cells[$"B{row}:D{row}"].Value = "TOTAL RESPIRATORY: ";

                    sheet.Cells[$"E{row}"].Formula = "=SUM(E86:E88)";
                    sheet.Cells[$"F{row}"].Formula = "=SUM(F86:F88)";
                    sheet.Cells[$"G{row}"].Formula = "=SUM(G86:G88)";
                    sheet.Cells[$"H{row}"].Formula = "=SUM(H86:H88)";


                    sheet.Cells[$"J{row}"].Formula = "=SUM(J86:J88)";
                    sheet.Cells[$"K{row}"].Formula = "=SUM(K86:K88)";
                    sheet.Cells[$"L{row}"].Formula = "=SUM(L86:L88)";
                    sheet.Cells[$"M{row}"].Formula = "=SUM(M86:M88)";


                    sheet.Cells[$"O{row}"].Formula = "=SUM(O86:O88)";
                    sheet.Cells[$"P{row}"].Formula = "=SUM(P86:P88)";
                    sheet.Cells[$"Q{row}"].Formula = "=SUM(Q86:Q88)";
                    sheet.Cells[$"R{row}"].Formula = "=SUM(R86:R88)";


                    sheet.Cells[$"T{row}"].Formula = "=SUM(T86:T88)";
                    sheet.Cells[$"U{row}"].Formula = "=SUM(U86:U88)";
                    sheet.Cells[$"V{row}"].Formula = "=SUM(V86:V88)";
                    sheet.Cells[$"W{row}"].Formula = "=SUM(W86:W88)";

                    sheet.Cells[$"I{row}"].Formula = "=SUM(I86:I88)";
                    sheet.Cells[$"N{row}"].Formula = "=SUM(N86:N88)";
                    sheet.Cells[$"S{row}"].Formula = "=SUM(S86:S88)";
                    sheet.Cells[$"X{row}"].Formula = "=SUM(X86:X88)";


                    sheet.Cells[$"Y{row}"].Formula = $"=SUM(Y86:Y88)";

                    row++;
                    ctrLabel++;
                }

                if (tally.IsHeader)
                {
                    specimenLabel = $"{ctrLabel}. {specimen}";
                    sheet.Cells[$"B{row}:D{row}"].Merge = true;
                    sheet.Cells[$"B{row}:D{row}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    sheet.Cells[$"B{row}:Y{row}"].Style.Fill.SetBackground(Color.Yellow);
                    sheet.Cells[$"B{row}:Y{row}"].Style.Font.Color.SetColor(Color.Red);
                    sheet.Cells[$"B{row}:D{row}"].Value = specimenLabel;
                    row++;

                    incrementFactor = 0;


                    continue;
                }
                else
                {
                    if (row >= 73 && row <= 75)
                    {
                        specimenLabel = $"{ctrLabel}. {specimen}";
                        sheet.Cells[$"B{row}:D{row}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    }
                    else
                    {
                        specimenLabel = specimen;
                        sheet.Cells[$"B{row}:D{row}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }
                }

                if(row == 90)
                {
                    sheet.Cells[$"B{row}:D{row}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    specimenLabel = $"{ctrLabel}. {specimen}";
                }

                sheet.Cells[$"B{row}:D{row}"].Merge = true;
                sheet.Cells[$"B{row}:D{row}"].Value = specimenLabel;

                sheet.Cells[$"E{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[1][2].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().IP);
                sheet.Cells[$"J{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[2][2].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().IP);
                sheet.Cells[$"O{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[3][2].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().IP);
                sheet.Cells[$"T{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[4][2].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().IP);

                sheet.Cells[$"F{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[1][2].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().ER);
                sheet.Cells[$"K{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[2][2].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().ER);
                sheet.Cells[$"P{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[3][2].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().ER);
                sheet.Cells[$"U{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[4][2].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().ER);

                sheet.Cells[$"G{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[1][2].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().MAGS);
                sheet.Cells[$"L{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[2][2].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().MAGS);
                sheet.Cells[$"Q{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[3][2].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().MAGS);
                sheet.Cells[$"V{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[4][2].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().MAGS);

                sheet.Cells[$"H{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[1][2].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().OPD);
                sheet.Cells[$"M{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[2][2].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().OPD);
                sheet.Cells[$"R{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[3][2].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().OPD);
                sheet.Cells[$"W{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[4][2].CategoryTallies.Where(x => x.SpecimenType == specimen).FirstOrDefault().OPD);

                sheet.Cells[$"I{row}"].Formula = $"=SUM(E{row},F{row},G{row},H{row})";
                sheet.Cells[$"N{row}"].Formula = $"=SUM(J{row},K{row},L{row},M{row})";
                sheet.Cells[$"S{row}"].Formula = $"=SUM(O{row},P{row},Q{row},R{row})";
                sheet.Cells[$"X{row}"].Formula = $"=SUM(T{row},U{row},V{row},W{row})";

                sheet.Cells[$"Y{row}"].Formula = $"=SUM(I{row},N{row},S{row},X{row})";

                row++;

                if(row == 91)
                {
                    sheet.Cells[$"B{row}:D{row}"].Merge = true;
                }

                ctrLabel += incrementFactor;
            }

            sheet.Cells[$"B{row}:D{row}"].Merge = true;
            row++;

            sheet.Cells["B92:Y92"].Style.Border.BorderAround(ExcelBorderStyle.Thick, Color.Black);
            sheet.Cells[$"B92:D92"].Merge = true;
            sheet.Cells[$"B92:D92"].Value = "TOTAL NO OF. CULTURE & SENSI RECEIVED:";
            sheet.Cells["B92:Y92"].Style.Border.BorderAround(ExcelBorderStyle.Thick, Color.Black);


            sheet.Cells[$"E92"].Formula = "=SUM(E71,E73:E75,E84,E89,E90)";
            sheet.Cells[$"F92"].Formula = "=SUM(F71,F73:F75,F84,F89,F90)";
            sheet.Cells["G92"].Formula = "=SUM(G71,G73:G75,G84,G89,G90)";
            sheet.Cells["H92"].Formula = "=SUM(H71,H73:H75,H84,H89,H90)";
            sheet.Cells["I92"].Formula = "=SUM(I71,I73:I75,I84,I89,I90)";

            sheet.Cells["J92"].Formula = "=SUM(J71,J73:J75,J84,J89,J90)";
            sheet.Cells["K92"].Formula = "=SUM(K71,K73:K75,K84,K89,K90)";
            sheet.Cells["L92"].Formula = "=SUM(L71,L73:L75,L84,L89,L90)";
            sheet.Cells["M92"].Formula = "=SUM(M71,M73:M75,M84,M89,M90)";
            sheet.Cells["N92"].Formula = "=SUM(N71,N73:N75,N84,N89,N90)";
            sheet.Cells["O92"].Formula = "=SUM(O71,O73:O75,O84,O89,O90)";
            sheet.Cells["P92"].Formula = "=SUM(P71,P73:P75,P84,P89,P90)";
            sheet.Cells["Q92"].Formula = "=SUM(Q71,Q73:Q75,Q84,Q89,Q90)";
            sheet.Cells["R92"].Formula = "=SUM(R71,R73:R75,R84,R89,R90)";
            sheet.Cells["S92"].Formula = "=SUM(S71,S73:S75,S84,S89,S90)";
            sheet.Cells["T92"].Formula = "=SUM(T71,T73:T75,T84,T89,T90)";
            sheet.Cells["U92"].Formula = "=SUM(U71,U73:U75,U84,U89,U90)";
            sheet.Cells["V92"].Formula = "=SUM(V71,V73:V75,V84,V89,V90)";
            sheet.Cells["W92"].Formula = "=SUM(W71,W73:W75,W84,W89,W90)";
            sheet.Cells["X92"].Formula = "=SUM(X71,X73:X75,X84,X89,X90)";
            sheet.Cells["Y92"].Formula = "=SUM(Y71,Y73:Y75,Y84,Y89,Y90)";

            sheet.Cells["I92"].Style.Font.Color.SetColor(Color.Red);
            sheet.Cells["N92"].Style.Font.Color.SetColor(Color.Red);
            sheet.Cells["S92"].Style.Font.Color.SetColor(Color.Red);
            sheet.Cells["X92"].Style.Font.Color.SetColor(Color.Red);
            sheet.Cells["Y92"].Style.Font.Color.SetColor(Color.Purple);


            #endregion

            #region Other Categories
            row = 98;

            tallies = monthlyReportEntity.TestTallies[1].Skip(3).Select(x=> new CategoryTally(x.CategoryName)).ToList();
            foreach (var tally in tallies)
            {
                var specimen = tally.SpecimenType;
                sheet.Cells[$"B{row}:D{row}"].Merge = true;

                sheet.Cells[$"B{row}:D{row}"].Value = $"{specimen}";

                sheet.Cells[$"E{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[1].Where(x => x.CategoryName == specimen).FirstOrDefault().CategoryTallies.FirstOrDefault().IP);
                sheet.Cells[$"J{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[2].Where(x => x.CategoryName == specimen).FirstOrDefault().CategoryTallies.FirstOrDefault().IP);
                sheet.Cells[$"O{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[3].Where(x => x.CategoryName == specimen).FirstOrDefault().CategoryTallies.FirstOrDefault().IP);
                sheet.Cells[$"T{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[4].Where(x => x.CategoryName == specimen).FirstOrDefault().CategoryTallies.FirstOrDefault().IP);

                sheet.Cells[$"F{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[1].Where(x => x.CategoryName == specimen).FirstOrDefault().CategoryTallies.FirstOrDefault().ER);
                sheet.Cells[$"K{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[2].Where(x => x.CategoryName == specimen).FirstOrDefault().CategoryTallies.FirstOrDefault().ER);
                sheet.Cells[$"P{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[3].Where(x => x.CategoryName == specimen).FirstOrDefault().CategoryTallies.FirstOrDefault().ER);
                sheet.Cells[$"U{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[4].Where(x => x.CategoryName == specimen).FirstOrDefault().CategoryTallies.FirstOrDefault().ER);

                sheet.Cells[$"G{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[1].Where(x => x.CategoryName == specimen).FirstOrDefault().CategoryTallies.FirstOrDefault().MAGS);
                sheet.Cells[$"L{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[2].Where(x => x.CategoryName == specimen).FirstOrDefault().CategoryTallies.FirstOrDefault().MAGS);
                sheet.Cells[$"Q{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[3].Where(x => x.CategoryName == specimen).FirstOrDefault().CategoryTallies.FirstOrDefault().MAGS);
                sheet.Cells[$"V{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[4].Where(x => x.CategoryName == specimen).FirstOrDefault().CategoryTallies.FirstOrDefault().MAGS);

                sheet.Cells[$"H{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[1].Where(x => x.CategoryName == specimen).FirstOrDefault().CategoryTallies.FirstOrDefault().OPD);
                sheet.Cells[$"M{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[2].Where(x => x.CategoryName == specimen).FirstOrDefault().CategoryTallies.FirstOrDefault().OPD);
                sheet.Cells[$"R{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[3].Where(x => x.CategoryName == specimen).FirstOrDefault().CategoryTallies.FirstOrDefault().OPD);
                sheet.Cells[$"W{row}"].Value = EmptyIfZero(monthlyReportEntity.TestTallies[4].Where(x => x.CategoryName == specimen).FirstOrDefault().CategoryTallies.FirstOrDefault().OPD);

                sheet.Cells[$"I{row}"].Formula = $"=SUM(E{row},F{row},G{row},H{row})";
                sheet.Cells[$"N{row}"].Formula = $"=SUM(J{row},K{row},L{row},M{row})";
                sheet.Cells[$"S{row}"].Formula = $"=SUM(O{row},P{row},Q{row},R{row})";
                sheet.Cells[$"X{row}"].Formula = $"=SUM(T{row},U{row},V{row},W{row})";

                sheet.Cells[$"Y{row}"].Formula = $"=SUM(I{row},N{row},S{row},X{row})";

                row++;
            }

            row++;
            sheet.Cells[$"B103:D103"].Merge = true;

            sheet.Cells[$"B103:D103"].Value = "TOTAL NO OF. SPECIMEN RECEIVED:";
            sheet.Cells["B103:Y103"].Style.Border.BorderAround(ExcelBorderStyle.Thick, Color.Black);

            sheet.Cells[$"E103"].Formula = "=SUM(E98:E102)";
            sheet.Cells["F103"].Formula = "=SUM(F98:F102)";
            sheet.Cells["G103"].Formula = "=SUM(G98:G102)";
            sheet.Cells["H103"].Formula = "=SUM(H98:H102)";
            sheet.Cells["I103"].Formula = "=SUM(I98:I102)";
            sheet.Cells["J103"].Formula = "=SUM(J98:J102)";
            sheet.Cells["K103"].Formula = "=SUM(K98:K102)";
            sheet.Cells["L103"].Formula = "=SUM(L98:L102)";
            sheet.Cells["M103"].Formula = "=SUM(M98:M102)";
            sheet.Cells["N103"].Formula = "=SUM(N98:N102)";
            sheet.Cells["O103"].Formula = "=SUM(O98:O102)";
            sheet.Cells["P103"].Formula = "=SUM(P98:P102)";
            sheet.Cells["Q103"].Formula = "=SUM(Q98:Q102)";
            sheet.Cells["R103"].Formula = "=SUM(R98:R102)";
            sheet.Cells["S103"].Formula = "=SUM(S98:S102)";
            sheet.Cells["T103"].Formula = "=SUM(T98:T102)";
            sheet.Cells["U103"].Formula = "=SUM(U98:U102)";
            sheet.Cells["V103"].Formula = "=SUM(V98:V102)";
            sheet.Cells["W103"].Formula = "=SUM(W98:W102)";
            sheet.Cells["X103"].Formula = "=SUM(X98:X102)";
            sheet.Cells["Y103"].Formula = "=SUM(Y98:Y102)";


            sheet.Cells["I98:I103"].Style.Font.Color.SetColor(Color.Red);
            sheet.Cells["N98:N103"].Style.Font.Color.SetColor(Color.Red);
            sheet.Cells["S98:S103"].Style.Font.Color.SetColor(Color.Red);
            sheet.Cells["X98:X103"].Style.Font.Color.SetColor(Color.Red);
            sheet.Cells["Y98:Y102"].Style.Font.Color.SetColor(Color.Blue);

            sheet.Cells["Y103"].Style.Font.Color.SetColor(Color.Purple);


            #endregion
            sheet.Calculate();
        }
        private static void InflateMonthlyReportRowContent(ExcelWorksheet sheet, MonthlyReportEntity monthlyReportEntity)
        {
            sheet.Cells["D1:H1"].Merge = true;
            sheet.Cells["D1:H1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Cells["D1:H1"].Style.Font.Bold = true;
            sheet.Cells["D1:H1"].Value = $"MONTH {DateTimeService.GetMonthFromCode(monthlyReportEntity.MonthCode)} {monthlyReportEntity.Year}";

            sheet.Cells["A2:C2"].Style.Font.Bold = true;
            sheet.Cells["A2:C2"].Merge = true;
            sheet.Cells["A2:C2"].Value = "BACTERIOLOGY";
            sheet.Cells["A2:C2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            sheet.Cells["A4:C4"].Merge = true;
            sheet.Cells["A4:C4"].Value = "Total Specimen Received";
            sheet.Cells["A4:C4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            sheet.Cells["A5:C5"].Merge = true;
            sheet.Cells["A5:C5"].Value = "Total Specimen Performed";
            sheet.Cells["A5:C5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;



            sheet.Cells["D4:H5,D8:H21,D25:H46,D49:H57,D61:H66"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            sheet.Cells["D4:H5,D8:H21,D25:H46,D49:H57,D61:H66"].Style.Border.Top.Color.SetColor(Color.Black);

            sheet.Cells["D4:H5,D8:H21,D25:H46,D49:H57,D61:H66"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            sheet.Cells["D4:H5,D8:H21,D25:H46,D49:H57,D61:H66"].Style.Border.Left.Color.SetColor(Color.Black);

            sheet.Cells["D4:H5,D8:H21,D25:H46,D49:H57,D61:H66"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            sheet.Cells["D4:H5,D8:H21,D25:H46,D49:H57,D61:H66"].Style.Border.Right.Color.SetColor(Color.Black);

            sheet.Cells["D4:H5,D8:H21,D25:H46,D49:H57,D61:H66"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            sheet.Cells["D4:H5,D8:H21,D25:H46,D49:H57,D61:H66"].Style.Border.Bottom.Color.SetColor(Color.Black);

            sheet.Cells["D3,D7,D24,D48,D60"].Value = "IP";
            sheet.Cells["E3,E7,E24,E48,E60"].Value = "ER";
            sheet.Cells["F3,F7,F24,F48,F60"].Value = "MAGS";
            sheet.Cells["G3,G7,G24,G48,G60"].Value = "OPD";
            sheet.Cells["H3,H7,H24,H48,H60"].Value = "TOTAL";


            sheet.Cells["D3:H68"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Cells["H3:H5,H7:H22,H24:H46,H48:H57,H60:H67"].Style.Font.Bold = true;

            sheet.Cells["D3:H3,D7:H7,D24:H24,D48:H48,D60:H60"].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);


            sheet.Cells["A7:C8"].Style.Font.Bold = true;
            sheet.Cells["A7:C8"].Merge = true;
            sheet.Cells["A7:C8"].Value = "AFB";
            sheet.Cells["A7:C8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            int row = 9;
            int refRow = 5;

            #region AFB
            var aggregateSpecimen = monthlyReportEntity.TestTallies[1][0].CategoryTallies.Select(x => x.SpecimenType).ToList();
            for (int i = 0; i < 13; i++)
            {
                sheet.Cells[$"A{row}:C{row}"].Merge = true;

                if (i == 12)
                {

                    refRow = 26;
                    sheet.Cells[$"A{row}:C{row}"].Value = $"{i + 1}. OTHERS";

                    sheet.Cells[$"D{row}"].Formula = $"=SUM(WEEKLY!E{refRow}, WEEKLY!J{refRow}, WEEKLY!O{refRow}, WEEKLY!T{refRow})";
                    sheet.Cells[$"E{row}"].Formula = $"=SUM(WEEKLY!F{refRow}, WEEKLY!K{refRow}, WEEKLY!P{refRow}, WEEKLY!U{refRow})";
                    sheet.Cells[$"F{row}"].Formula = $"=SUM(WEEKLY!G{refRow}, WEEKLY!L{refRow}, WEEKLY!Q{refRow}, WEEKLY!V{refRow})";
                    sheet.Cells[$"G{row}"].Formula = $"=SUM(WEEKLY!H{refRow}, WEEKLY!M{refRow}, WEEKLY!R{refRow}, WEEKLY!W{refRow})";
                    sheet.Cells[$"H{row}"].Formula = $"=SUM(D{row}:G{row})";

                    row++;

                    sheet.Cells[$"A{row}:C{row}"].Merge = true;
                    sheet.Cells[$"A{row}:C{row}"].Value = "TOTAL";
                    sheet.Cells[$"A{row}:C{row}"].Style.Font.Bold = true;

                    sheet.Cells[$"D{row}"].Formula = $"=SUM(D9:D21)";
                    sheet.Cells[$"E{row}"].Formula = $"=SUM(E9:E21)";
                    sheet.Cells[$"F{row}"].Formula = $"=SUM(F9:F21)";
                    sheet.Cells[$"G{row}"].Formula = $"=SUM(G9:G21)";
                    sheet.Cells[$"H{row}"].Formula = $"=SUM(H9:H21)";
                    sheet.Cells[$"H{row}"].Style.Font.Color.SetColor(Color.Red);
                    sheet.Cells[$"H{row}"].Style.Font.Bold = true;

                    sheet.Cells[$"A{row}:H{row}"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[$"A{row}:H{row}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[$"A{row}:H{row}"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[$"A{row}:H{row}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[$"A{row}:H{row}"].Style.Border.Top.Color.SetColor(Color.Black);
                    sheet.Cells[$"A{row}:H{row}"].Style.Border.Bottom.Color.SetColor(Color.Black);
                    sheet.Cells[$"A{row}:H{row}"].Style.Border.Left.Color.SetColor(Color.Black);
                    sheet.Cells[$"A{row}:H{row}"].Style.Border.Right.Color.SetColor(Color.Black);

                    sheet.Cells[$"A{row}:C{row}"].Value = $"TOTAL";
                    sheet.Cells[$"A{row}:C{row}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;



                }
                else
                {
                    sheet.Cells[$"A{row}:C{row}"].Value = $"{i + 1}. {aggregateSpecimen[i]}";

                    sheet.Cells[$"D{row}"].Formula = $"=SUM(WEEKLY!E{refRow}, WEEKLY!J{refRow}, WEEKLY!O{refRow}, WEEKLY!T{refRow})";
                    sheet.Cells[$"E{row}"].Formula = $"=SUM(WEEKLY!F{refRow}, WEEKLY!K{refRow}, WEEKLY!P{refRow}, WEEKLY!U{refRow})";
                    sheet.Cells[$"F{row}"].Formula = $"=SUM(WEEKLY!G{refRow}, WEEKLY!L{refRow}, WEEKLY!Q{refRow}, WEEKLY!V{refRow})";
                    sheet.Cells[$"G{row}"].Formula = $"=SUM(WEEKLY!H{refRow}, WEEKLY!M{refRow}, WEEKLY!R{refRow}, WEEKLY!W{refRow})";
                    sheet.Cells[$"H{row}"].Formula = $"=SUM(D{row}:G{row})";

                }

                refRow++;
                row++;
            }

            #endregion

            #region GS
            row = 26;
            refRow = 34;

            sheet.Cells[$"A24:C25"].Merge = true;
            sheet.Cells[$"A24:C25"].Value = "GRAM STAINING";
            sheet.Cells[$"A24:C24"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Cells[$"A24:C24"].Style.Font.Bold = true;


            aggregateSpecimen = monthlyReportEntity.TestTallies[1][1].CategoryTallies.Select(x => x.SpecimenType).ToList();
            for (int i = 0; i < 20; i++)
            {
                sheet.Cells[$"A{row}:C{row}"].Merge = true;


                sheet.Cells[$"A{row}:C{row}"].Value = $"{i + 1}. {aggregateSpecimen[i]}";

                sheet.Cells[$"D{row}"].Formula = $"=SUM(WEEKLY!E{refRow}, WEEKLY!J{refRow}, WEEKLY!O{refRow}, WEEKLY!T{refRow})";
                sheet.Cells[$"E{row}"].Formula = $"=SUM(WEEKLY!F{refRow}, WEEKLY!K{refRow}, WEEKLY!P{refRow}, WEEKLY!U{refRow})";
                sheet.Cells[$"F{row}"].Formula = $"=SUM(WEEKLY!G{refRow}, WEEKLY!L{refRow}, WEEKLY!Q{refRow}, WEEKLY!V{refRow})";
                sheet.Cells[$"G{row}"].Formula = $"=SUM(WEEKLY!H{refRow}, WEEKLY!M{refRow}, WEEKLY!R{refRow}, WEEKLY!W{refRow})";
                sheet.Cells[$"H{row}"].Formula = $"=SUM(D{row}:G{row})";

                if (i == 19)
                {

                    row++;
                    sheet.Cells[$"D{row}"].Formula = $"=SUM(D26:D45)";
                    sheet.Cells[$"E{row}"].Formula = $"=SUM(E26:E45)";
                    sheet.Cells[$"F{row}"].Formula = $"=SUM(F26:F45)";
                    sheet.Cells[$"G{row}"].Formula = $"=SUM(G26:G45)";
                    sheet.Cells[$"H{row}"].Formula = $"=SUM(H26:H45)";
                    sheet.Cells[$"H{row}"].Style.Font.Color.SetColor(Color.Red);


                    sheet.Cells[$"H{row}"].Style.Font.Color.SetColor(Color.Red);
                    sheet.Cells[$"H{row}"].Style.Font.Bold = true;

                    sheet.Cells[$"A{row}:H{row}"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[$"A{row}:H{row}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[$"A{row}:H{row}"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[$"A{row}:H{row}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[$"A{row}:H{row}"].Style.Border.Top.Color.SetColor(Color.Black);
                    sheet.Cells[$"A{row}:H{row}"].Style.Border.Bottom.Color.SetColor(Color.Black);
                    sheet.Cells[$"A{row}:H{row}"].Style.Border.Left.Color.SetColor(Color.Black);
                    sheet.Cells[$"A{row}:H{row}"].Style.Border.Right.Color.SetColor(Color.Black);

                    sheet.Cells[$"A{row}:C{row}"].Merge = true;
                    sheet.Cells[$"A{row}:C{row}"].Value = $"TOTAL";
                    sheet.Cells[$"A{row}:C{row}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                }

                refRow++;
                row++;
            }

            #endregion

            #region C/S

            row = 50;

            sheet.Cells[$"A48:C49"].Merge = true;
            sheet.Cells[$"A48:C49"].Value = "CULTURE AND SENSITIVITY";
            sheet.Cells[$"A48:C49"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Cells[$"A48:C49"].Style.Font.Bold = true;

            for (int i = 0; i < 7; i++)
            {
                sheet.Cells[$"A{row}:C{row}"].Merge = true;
                string specimenLabel = string.Empty;

                switch (i + 1)
                {
                    case 1:
                        refRow = 71;
                        specimenLabel = "CSF,AF,DF";
                        break;
                    case 2:
                        refRow = 73;
                        specimenLabel = "RECTAL SWAB/STOOL";
                        break;
                    case 3:
                        refRow = 74;
                        specimenLabel = "URINE";
                        break;
                    case 4:
                        refRow = 75;
                        specimenLabel = "BLOOD";
                        break;
                    case 5:
                        refRow = 84;
                        specimenLabel = "WOUND/ABSCESS";
                        break;
                    case 6:
                        refRow = 89;
                        specimenLabel = "RESPIRATORY DISCHARGE";
                        break;
                    case 7:
                        refRow = 90;
                        specimenLabel = "ENVIRONMENTAL CULTURE";
                        break;
                }

                sheet.Cells[$"A{row}:C{row}"].Value = $"{i + 1}. {specimenLabel}";


                sheet.Cells[$"D{row}"].Formula = $"=SUM(WEEKLY!E{refRow}, WEEKLY!J{refRow}, WEEKLY!O{refRow}, WEEKLY!T{refRow})";
                sheet.Cells[$"E{row}"].Formula = $"=SUM(WEEKLY!F{refRow}, WEEKLY!K{refRow}, WEEKLY!P{refRow}, WEEKLY!U{refRow})";
                sheet.Cells[$"F{row}"].Formula = $"=SUM(WEEKLY!G{refRow}, WEEKLY!L{refRow}, WEEKLY!Q{refRow}, WEEKLY!V{refRow})";
                sheet.Cells[$"G{row}"].Formula = $"=SUM(WEEKLY!H{refRow}, WEEKLY!M{refRow}, WEEKLY!R{refRow}, WEEKLY!W{refRow})";
                sheet.Cells[$"H{row}"].Formula = $"=SUM(D{row}:G{row})";


                refRow++;
                row++;
            }

            sheet.Cells[$"A{row}:C{row}"].Merge = true;
            sheet.Cells[$"D{row}"].Formula = $"=SUM(D50:D56)";
            sheet.Cells[$"E{row}"].Formula = $"=SUM(E50:E56)";
            sheet.Cells[$"F{row}"].Formula = $"=SUM(F50:F56)";
            sheet.Cells[$"G{row}"].Formula = $"=SUM(G50:G56)";
            sheet.Cells[$"H{row}"].Formula = $"=SUM(H50:H56)";

            sheet.Cells[$"H{row}"].Style.Font.Color.SetColor(Color.Red);
            sheet.Cells[$"H{row}"].Style.Font.Bold = true;
            sheet.Cells[$"A{row}:H{row}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            sheet.Cells[$"A{row}:C{row}"].Value = $"TOTAL";
            sheet.Cells[$"A{row}:C{row}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            #endregion

            #region Other Categories 


            row = 61;
            refRow = 98;

            sheet.Cells[$"A59:C60"].Merge = true;
            sheet.Cells[$"A59:C60"].Value = "";
            sheet.Cells[$"A59:C60"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            aggregateSpecimen = monthlyReportEntity.TestTallies[1].Skip(3).Select(x => x.CategoryName).ToList();
            for (int i = 0; i < aggregateSpecimen.Count; i++)
            {
                sheet.Cells[$"A{row}:C{row}"].Merge = true;
                sheet.Cells[$"A{row}:C{row}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells[$"A{row}:C{row}"].Style.Font.Bold = true;

                sheet.Cells[$"A{row}:C{row}"].Value = $"{aggregateSpecimen[i]}";

                sheet.Cells[$"D{row}"].Formula = $"=SUM(WEEKLY!E{refRow}, WEEKLY!J{refRow}, WEEKLY!O{refRow}, WEEKLY!T{refRow})";
                sheet.Cells[$"E{row}"].Formula = $"=SUM(WEEKLY!F{refRow}, WEEKLY!K{refRow}, WEEKLY!P{refRow}, WEEKLY!U{refRow})";
                sheet.Cells[$"F{row}"].Formula = $"=SUM(WEEKLY!G{refRow}, WEEKLY!L{refRow}, WEEKLY!Q{refRow}, WEEKLY!V{refRow})";
                sheet.Cells[$"G{row}"].Formula = $"=SUM(WEEKLY!H{refRow}, WEEKLY!M{refRow}, WEEKLY!R{refRow}, WEEKLY!W{refRow})";
                sheet.Cells[$"H{row}"].Formula = $"=SUM(D{row}:G{row})";

                refRow++;
                row++;
            }

            sheet.Cells[$"A{row}:C{row}"].Merge = true;
            sheet.Cells[$"D{row}"].Formula = $"=SUM(D61:D65)";
            sheet.Cells[$"E{row}"].Formula = $"=SUM(E61:E65)";
            sheet.Cells[$"F{row}"].Formula = $"=SUM(F61:F65)";
            sheet.Cells[$"G{row}"].Formula = $"=SUM(G61:G65)";
            sheet.Cells[$"H{row}"].Formula = $"=SUM(H61:H65)";
            sheet.Cells[$"H{row}"].Style.Font.Color.SetColor(Color.Red);
            sheet.Cells[$"H{row}"].Style.Font.Bold = true;
            sheet.Cells[$"A{row}:H{row}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            sheet.Cells[$"A{row}:C{row}"].Value = $"TOTAL";
            sheet.Cells[$"A{row}:C{row}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            #endregion  

            sheet.Cells["J1:J100"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Cells["J3"].Value = "TAT";

            sheet.Cells["J8"].Value = "TAT AFB";
            sheet.Cells["I9"].Value="IP";
            sheet.Cells["I10"].Value = "ER";
            sheet.Cells["I11"].Value = "MAGS";
            sheet.Cells["I12"].Value = "OPD";


            sheet.Cells["J25"].Value = "TAT GS";
            sheet.Cells["I26"].Value = "IP";
            sheet.Cells["I27"].Value = "ER";
            sheet.Cells["I28"].Value = "MAGS";
            sheet.Cells["I29"].Value = "OPD";


            sheet.Cells["J49"].Value = "TAT CS";
            sheet.Cells["I50"].Value = "IP";
            sheet.Cells["I51"].Value = "ER";
            sheet.Cells["I52"].Value = "MAGS";
            sheet.Cells["I53"].Value = "OPD";


            sheet.Cells["J60"].Value = "TAT HUMAN MILK CULTURE";
            sheet.Cells["I61"].Value = "IP";

            sheet.Cells["J62"].Value = "TAT GEN X";
            sheet.Cells["I63"].Value = "IP";
            sheet.Cells["I64"].Value = "ER";
            sheet.Cells["I65"].Value = "MAGS";
            sheet.Cells["I66"].Value = "OPD";

            sheet.Cells["I1:I100"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;



            sheet.Cells["J68"].Value = "TAT KOH";
            sheet.Cells["I69"].Value = "IP";
            sheet.Cells["I70"].Value = "ER";
            sheet.Cells["I71"].Value = "MAGS";
            sheet.Cells["I72"].Value = "OPD";



            sheet.Cells["J73"].Value = "TAT INDIA INK";
            sheet.Cells["I74"].Value = "IP";
            sheet.Cells["I75"].Value = "ER";
            sheet.Cells["I76"].Value = "MAGS";
            sheet.Cells["I77"].Value = "OPD";



            sheet.Cells["J78"].Value = "TAT XDR";
            sheet.Cells["I79"].Value = "IP";
            sheet.Cells["I80"].Value = "ER";
            sheet.Cells["I81"].Value = "MAGS";
            sheet.Cells["I82"].Value = "OPD";

            sheet.Cells["J3"].Style.Font.Bold = true;
            sheet.Cells["J8"].Style.Font.Bold = true;
            sheet.Cells["J25"].Style.Font.Bold = true;
            sheet.Cells["J49"].Style.Font.Bold = true;
            sheet.Cells["J60"].Style.Font.Bold = true;
            sheet.Cells["J62"].Style.Font.Bold = true;
            sheet.Cells["J68"].Style.Font.Bold = true;
            sheet.Cells["J73"].Style.Font.Bold = true;
            sheet.Cells["J78"].Style.Font.Bold = true;


            sheet.Calculate();
        }

        private static int? EmptyIfZero(int num)
        {
            return (num == 0) ? (int?)null : num;
        }

        private static int ZeroIfEmpty(string val)
        {
            if (string.IsNullOrEmpty(val)) return 0;
            return Convert.ToInt32(val);
        }

    }
}

using MonthlyReportEntrySoftware.Entities;
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

            var bannerTitleRange = sheet.Cells[$"E1:W1,E29:W29,E50:W50,E81:W81"];
            bannerTitleRange.Merge = true;
            bannerTitleRange.Value = $"WEEKLY REPORT {DateTimeService.GetMonthFromCode(monthlyReportEntity.MonthCode)} {monthlyReportEntity.Year}";
            bannerTitleRange.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            bannerTitleRange.Style.Font.Size = 20;
            bannerTitleRange.Style.Fill.SetBackground(BannerTitlecolor);


            sheet.Cells["E3,J3,O3,T3,E31,J31,O31,T31,E52,J52,O52,T52,E83,J83,O83,T83"].Value = "IP";
            sheet.Cells["E3:E25,J3:J25,O3:O25,T3:T25,E31:E46,J31:J46,O31:O46,T31:T46,E52:E77,J52:J77,O52:O77,T52:T77,E83:E88,J83:J88,O83:O88,T83:T88"].Style.Fill.SetBackground(IPColumnColor);

            sheet.Cells["F3,K3,P3,U3,F31,K31,P31,U31,F52,K52,P52,U52,F83,K83,P83,U83"].Value = "ER";
            sheet.Cells["F3:F25,K3:K25,P3:P25,U3:U25,F31:F46,K31:K46,P31:P46,U31:U46,F52:F77,K52:K77,P52:P77,U52:U77,F83:F88,K83:K88,P83:P88,U83:U88"].Style.Fill.SetBackground(ERColumnColor);

            sheet.Cells["G3,L3,Q3,V3,G31,L31,Q31,V31,G52,L52,Q52,V52,G83,L83,Q83,V83"].Value = "MAGS";
            sheet.Cells["G3:G25,L3:L25,Q3:Q25,V3:V25,G31:G46,L31:L46,Q31:Q46,V31:V46,G52:G77,L52:L77,Q52:Q77,V52:V77,G83:G88,L83:L88,Q83:Q88,V83:V88"].Style.Fill.SetBackground(MAGSColumnColor);

            sheet.Cells["H3,M3,R3,W3,H31,M31,R31,W31,H52,M52,R52,W52,H83,M83,R83,W83"].Value = "OP";
            sheet.Cells["H3:H25,M3:M25,R3:R25,W3:W25,H31:H46,M31:M46,R31:R46,W31:W46,H52:H77,M52:M77,R52:R77,W52:W77,H83:H88,M83:M88,R83:R88,W83:W88"].Style.Fill.SetBackground(OPColumnColor);


            sheet.Cells["B2:Y26,B30:Y47,B51:Y78,B82:Y89"].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
            sheet.Cells["B2:D3"].Merge = true;
            sheet.Cells["B2:D3"].Value = "BACTERIOLOGY";
            sheet.Cells["B2:D3"].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
            sheet.Cells["B2:D3"].Style.Font.Bold = true;
            sheet.Cells["B2:D3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Cells["B2:D3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            sheet.Cells["B30:D31"].Merge = true;
            sheet.Cells["B30:D31"].Value = "GRAM STAINING";
            sheet.Cells["B30:D31"].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
            sheet.Cells["B30:D31"].Style.Font.Bold = true;
            sheet.Cells["B30:D31"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Cells["B30:D31"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            sheet.Cells["B51:D52"].Merge = true;
            sheet.Cells["B51:D52"].Value = "CULTURE AND SENSI";
            sheet.Cells["B51:D52"].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
            sheet.Cells["B51:D52"].Style.Font.Bold = true;
            sheet.Cells["B51:D52"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Cells["B51:D52"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            sheet.Cells["B82:D83"].Merge = true;
            sheet.Cells["B82:D83"].Value = "OTHER TESTS";
            sheet.Cells["B82:D83"].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
            sheet.Cells["B82:D83"].Style.Font.Bold = true;
            sheet.Cells["B82:D83"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Cells["B82:D83"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            var tempRange = sheet.Cells["E2:H2,E30:H30,E51:H51,E82:H82"];
            tempRange.Value = "1ST WK";
            tempRange.Merge = true;
            tempRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            tempRange.Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);

            tempRange = sheet.Cells["J2:M2,J30:M30,J51:M51,J82:M82"];
            tempRange.Value = "2ND WK";
            tempRange.Merge = true;
            tempRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            tempRange.Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
            tempRange.Style.Fill.SetBackground(ERColumnColor);

            tempRange = sheet.Cells["O2:R2,O30:R30,O51:R51,O82:R82"];
            tempRange.Value = "3RD WK";
            tempRange.Merge = true;
            tempRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            tempRange.Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);

            tempRange = sheet.Cells["T2:W2,T30:W30,T51:W51,T82:W82"];
            tempRange.Value = "4TH WK";
            tempRange.Merge = true;
            tempRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            tempRange.Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);

            sheet.Cells["Y2,Y30,Y51,Y83"].Value = "TOTAL";


            sheet.Cells[$"B24:H24,J24:M24,O24:R24,T24:W24"].Style.Fill.SetBackground(Color.Gold);
            sheet.Cells[$"B53:Y53,B66:Y66,B71:Y71"].Style.Fill.SetBackground(Color.Yellow);
            sheet.Cells[$"B61:Y61,B70:Y70,B74:Y74"].Style.Fill.SetBackground(Color.White);
            sheet.Cells[$"B53:Y53,B66:Y66,B71:Y71"].Style.Font.Color.SetColor(Color.Red);
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

                    for (int i = 5; i <= 23; i++)
                    {

                        if (i == 16 || i == 17) continue;
                        string currentSpecimen = sheet.Cells[$"C{i}"].Text;
                        int _index = ((currentSpecimen.IndexOf(" ") < 0) ? 0 : currentSpecimen.IndexOf(" "));
                        currentSpecimen = currentSpecimen.Substring(_index).Trim();
                        
                        monthlyReportEntity.TestTallies[1][0].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().IP = ZeroIfEmpty(sheet.Cells[$"E{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[2][0].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().IP = ZeroIfEmpty(sheet.Cells[$"J{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[3][0].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().IP = ZeroIfEmpty(sheet.Cells[$"O{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[4][0].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().IP = ZeroIfEmpty(sheet.Cells[$"T{i}"].Value?.ToString());

                        monthlyReportEntity.TestTallies[1][0].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().ER = ZeroIfEmpty(sheet.Cells[$"F{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[2][0].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().ER = ZeroIfEmpty(sheet.Cells[$"K{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[3][0].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().ER = ZeroIfEmpty(sheet.Cells[$"P{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[4][0].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().ER = ZeroIfEmpty(sheet.Cells[$"U{i}"].Value?.ToString());

                        monthlyReportEntity.TestTallies[1][0].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().MAGS = ZeroIfEmpty(sheet.Cells[$"G{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[2][0].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().MAGS = ZeroIfEmpty(sheet.Cells[$"L{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[3][0].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().MAGS = ZeroIfEmpty(sheet.Cells[$"Q{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[4][0].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().MAGS = ZeroIfEmpty(sheet.Cells[$"V{i}"].Value?.ToString());

                        monthlyReportEntity.TestTallies[1][0].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().OPD = ZeroIfEmpty(sheet.Cells[$"H{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[2][0].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().OPD = ZeroIfEmpty(sheet.Cells[$"M{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[3][0].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().OPD = ZeroIfEmpty(sheet.Cells[$"R{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[4][0].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().OPD = ZeroIfEmpty(sheet.Cells[$"W{i}"].Value?.ToString());

                    }

                    #endregion

                    #region GS

                    for (int i = 32; i <= 45; i++)
                    {

                        string currentSpecimen = sheet.Cells[$"C{i}"].Text;
                        int _index = ((currentSpecimen.IndexOf(" ") < 0) ? 0 : currentSpecimen.IndexOf(" "));
                        currentSpecimen = currentSpecimen.Substring(_index).Trim();

                        monthlyReportEntity.TestTallies[1][1].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().IP = ZeroIfEmpty(sheet.Cells[$"E{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[2][1].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().IP = ZeroIfEmpty(sheet.Cells[$"J{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[3][1].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().IP = ZeroIfEmpty(sheet.Cells[$"O{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[4][1].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().IP = ZeroIfEmpty(sheet.Cells[$"T{i}"].Value?.ToString());

                        monthlyReportEntity.TestTallies[1][1].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().ER = ZeroIfEmpty(sheet.Cells[$"F{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[2][1].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().ER = ZeroIfEmpty(sheet.Cells[$"K{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[3][1].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().ER = ZeroIfEmpty(sheet.Cells[$"P{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[4][1].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().ER = ZeroIfEmpty(sheet.Cells[$"U{i}"].Value?.ToString());

                        monthlyReportEntity.TestTallies[1][1].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().MAGS = ZeroIfEmpty(sheet.Cells[$"G{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[2][1].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().MAGS = ZeroIfEmpty(sheet.Cells[$"L{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[3][1].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().MAGS = ZeroIfEmpty(sheet.Cells[$"Q{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[4][1].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().MAGS = ZeroIfEmpty(sheet.Cells[$"V{i}"].Value?.ToString());

                        monthlyReportEntity.TestTallies[1][1].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().OPD = ZeroIfEmpty(sheet.Cells[$"H{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[2][1].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().OPD = ZeroIfEmpty(sheet.Cells[$"M{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[3][1].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().OPD = ZeroIfEmpty(sheet.Cells[$"R{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[4][1].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().OPD = ZeroIfEmpty(sheet.Cells[$"W{i}"].Value?.ToString());

                    }

                    #endregion

                    #region CS

                    for (int i = 53; i <= 76; i++)
                    {
                        if (i == 53 || i == 61 || i == 62 || i == 66 || i == 70 || i == 71 || i == 74 || i == 75) continue;
                        string currentSpecimen = sheet.Cells[$"C{i}"].Text;
                        int _index = ((currentSpecimen.IndexOf(" ") < 0) ? 0 : currentSpecimen.IndexOf(" "));
                        if (i == 59 || i == 60 || i == 69 || i == 76) _index = 0;
                        currentSpecimen = currentSpecimen.Substring(_index).Trim();

                        monthlyReportEntity.TestTallies[1][2].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().IP = ZeroIfEmpty(sheet.Cells[$"E{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[2][2].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().IP = ZeroIfEmpty(sheet.Cells[$"J{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[3][2].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().IP = ZeroIfEmpty(sheet.Cells[$"O{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[4][2].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().IP = ZeroIfEmpty(sheet.Cells[$"T{i}"].Value?.ToString());

                        monthlyReportEntity.TestTallies[1][2].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().ER = ZeroIfEmpty(sheet.Cells[$"F{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[2][2].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().ER = ZeroIfEmpty(sheet.Cells[$"K{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[3][2].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().ER = ZeroIfEmpty(sheet.Cells[$"P{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[4][2].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().ER = ZeroIfEmpty(sheet.Cells[$"U{i}"].Value?.ToString());

                        monthlyReportEntity.TestTallies[1][2].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().MAGS = ZeroIfEmpty(sheet.Cells[$"G{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[2][2].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().MAGS = ZeroIfEmpty(sheet.Cells[$"L{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[3][2].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().MAGS = ZeroIfEmpty(sheet.Cells[$"Q{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[4][2].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().MAGS = ZeroIfEmpty(sheet.Cells[$"V{i}"].Value?.ToString());

                        monthlyReportEntity.TestTallies[1][2].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().OPD = ZeroIfEmpty(sheet.Cells[$"H{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[2][2].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().OPD = ZeroIfEmpty(sheet.Cells[$"M{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[3][2].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().OPD = ZeroIfEmpty(sheet.Cells[$"R{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[4][2].CategoryTallies.Where(x => x.SpecimenType == currentSpecimen.Trim()).FirstOrDefault().OPD = ZeroIfEmpty(sheet.Cells[$"W{i}"].Value?.ToString());

                    }

                    #endregion

                    #region Other Tests

                    for (int i = 84; i <= 87; i++)
                    {
                        string currentSpecimen = sheet.Cells[$"C{i}"].Text;

                        monthlyReportEntity.TestTallies[1][i - 81].CategoryTallies.FirstOrDefault().IP = ZeroIfEmpty(sheet.Cells[$"E{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[2][i - 81].CategoryTallies.FirstOrDefault().IP = ZeroIfEmpty(sheet.Cells[$"J{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[3][i - 81].CategoryTallies.FirstOrDefault().IP = ZeroIfEmpty(sheet.Cells[$"O{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[4][i - 81].CategoryTallies.FirstOrDefault().IP = ZeroIfEmpty(sheet.Cells[$"T{i}"].Value?.ToString());

                        monthlyReportEntity.TestTallies[1][i - 81].CategoryTallies.FirstOrDefault().ER = ZeroIfEmpty(sheet.Cells[$"F{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[2][i - 81].CategoryTallies.FirstOrDefault().ER = ZeroIfEmpty(sheet.Cells[$"K{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[3][i - 81].CategoryTallies.FirstOrDefault().ER = ZeroIfEmpty(sheet.Cells[$"P{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[4][i - 81].CategoryTallies.FirstOrDefault().ER = ZeroIfEmpty(sheet.Cells[$"U{i}"].Value?.ToString());

                        monthlyReportEntity.TestTallies[1][i - 81].CategoryTallies.FirstOrDefault().MAGS = ZeroIfEmpty(sheet.Cells[$"G{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[2][i - 81].CategoryTallies.FirstOrDefault().MAGS = ZeroIfEmpty(sheet.Cells[$"L{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[3][i - 81].CategoryTallies.FirstOrDefault().MAGS = ZeroIfEmpty(sheet.Cells[$"Q{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[4][i - 81].CategoryTallies.FirstOrDefault().MAGS = ZeroIfEmpty(sheet.Cells[$"V{i}"].Value?.ToString());

                        monthlyReportEntity.TestTallies[1][i - 81].CategoryTallies.FirstOrDefault().OPD = ZeroIfEmpty(sheet.Cells[$"H{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[2][i - 81].CategoryTallies.FirstOrDefault().OPD = ZeroIfEmpty(sheet.Cells[$"M{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[3][i - 81].CategoryTallies.FirstOrDefault().OPD = ZeroIfEmpty(sheet.Cells[$"R{i}"].Value?.ToString());
                        monthlyReportEntity.TestTallies[4][i - 81].CategoryTallies.FirstOrDefault().OPD = ZeroIfEmpty(sheet.Cells[$"W{i}"].Value?.ToString());

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
            sheet.Cells[$"B1:Y89"].Style.Font.Bold = true;
            sheet.Cells[$"B1:Y89"].Style.Font.Size = 10;

            sheet.Cells["I5:I15,I32:I45,I61:I65,I70,I74,I84:I87"].Style.Font.Color.SetColor(Color.Red);
            sheet.Cells["N5:N15,N32:N45,N61:N65,N70,N74,N84:N87"].Style.Font.Color.SetColor(Color.Red);
            sheet.Cells["S5:S15,S32:S45,S61:S65,S70,S74,S84:S87"].Style.Font.Color.SetColor(Color.Red);
            sheet.Cells["X5:X15,X32:X45,X61:X65,X70,X74,X84:X87"].Style.Font.Color.SetColor(Color.Red);
            sheet.Cells["Y54:Y60,Y67:Y69,Y72:Y73"].Style.Font.Color.SetColor(Color.Red);
            sheet.Cells["Y5:Y15,Y32:Y45,Y61:Y65,Y70,Y76,Y74,Y84:Y88"].Style.Font.Color.SetColor(Color.Blue);

            sheet.Cells["B2:Y26,B30:Y47,B51:Y78,B82:Y89"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            sheet.Cells["B2:Y26,B30:Y47,B51:Y78,B82:Y89"].Style.Border.Top.Color.SetColor(Color.Black);

            sheet.Cells["B2:Y26,B30:Y47,B51:Y78,B82:Y89"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            sheet.Cells["B2:Y26,B30:Y47,B51:Y78,B82:Y89"].Style.Border.Bottom.Color.SetColor(Color.Black);

            sheet.Cells["B2:Y26,B30:Y47,B51:Y78,B82:Y89"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            sheet.Cells["B2:Y26,B30:Y47,B51:Y78,B82:Y89"].Style.Border.Left.Color.SetColor(Color.Black);

            sheet.Cells["B2:Y26,B30:Y47,B51:Y78,B82:Y89"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            sheet.Cells["B2:Y26,B30:Y47,B51:Y78,B82:Y89"].Style.Border.Right.Color.SetColor(Color.Black);


            var specimens = monthlyReportEntity.TestTallies[1][0].CategoryTallies.Select(x => x.SpecimenType).ToList();

            #region AFB Category
            foreach (var specimen in specimens)
            {
                string counterLabel = (row - 4).ToString() + ". ";
                sheet.Cells[$"B{row}:D{row}"].Merge = true;
                if (row >= 26) break;


                if (row - 4 == 12)
                {
                    sheet.Cells[$"Y{row}"].Formula = $"=SUM(Y5:Y15)";
                    sheet.Cells[$"Y{row}"].Style.Font.Color.SetColor(Color.Red);

                    row++;
                    sheet.Cells[$"B{row}:D{row}"].Merge = true;
                    sheet.Cells[$"B{row}:D{row}"].Value = "OTHERS:";
                    sheet.Cells[$"E{row}:W{row}"].Style.Fill.SetBackground(Color.White);

                    row++;
                    sheet.Cells[$"B{row}:D{row}"].Merge = true;
                }


                if (row >= 18)
                {
                    counterLabel = string.Empty;
                    sheet.Cells[$"B{row}:D{row}"].Style.Fill.SetBackground((row == 24) ? Color.Gold : GrayColor);
                }

                sheet.Cells[$"B{row}:D{row}"].Value = $"{counterLabel} {specimen}";

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

            sheet.Cells[$"E{row}"].Formula = "=SUM(E18:E23)";
            sheet.Cells[$"F{row}"].Formula = "=SUM(F18:F23)";
            sheet.Cells[$"G{row}"].Formula = "=SUM(G18:G23)";
            sheet.Cells[$"H{row}"].Formula = "=SUM(H18:H23)";


            sheet.Cells[$"J{row}"].Formula = "=SUM(J18:J23)";
            sheet.Cells[$"K{row}"].Formula = "=SUM(K18:K23)";
            sheet.Cells[$"L{row}"].Formula = "=SUM(L18:L23)";
            sheet.Cells[$"M{row}"].Formula = "=SUM(M18:M23)";


            sheet.Cells[$"O{row}"].Formula = "=SUM(O18:O23)";
            sheet.Cells[$"P{row}"].Formula = "=SUM(P18:P23)";
            sheet.Cells[$"Q{row}"].Formula = "=SUM(Q18:Q23)";
            sheet.Cells[$"R{row}"].Formula = "=SUM(R18:R23)";


            sheet.Cells[$"T{row}"].Formula = "=SUM(T18:T23)";
            sheet.Cells[$"U{row}"].Formula = "=SUM(U18:U23)";
            sheet.Cells[$"V{row}"].Formula = "=SUM(V18:V23)";
            sheet.Cells[$"W{row}"].Formula = "=SUM(W18:W23)";

            sheet.Cells[$"I{row}"].Formula = "=SUM(I18:I23)";
            sheet.Cells[$"N{row}"].Formula = "=SUM(N18:N23)";
            sheet.Cells[$"S{row}"].Formula = "=SUM(S18:S23)";
            sheet.Cells[$"X{row}"].Formula = "=SUM(X18:X23)";

            sheet.Cells[$"Y{row}"].Formula = "=SUM(I24,N24,S24,X24)";

            sheet.Cells[$"I{row}"].Style.Font.Color.SetColor(Color.Red);
            sheet.Cells[$"N{row}"].Style.Font.Color.SetColor(Color.Red);
            sheet.Cells[$"S{row}"].Style.Font.Color.SetColor(Color.Red);
            sheet.Cells[$"X{row}"].Style.Font.Color.SetColor(Color.Red);
            sheet.Cells[$"Y{row}"].Style.Font.Color.SetColor(Color.Blue);

            sheet.Cells["B25:D25"].Merge = true;

            sheet.Cells["B26:D26"].Merge = true;
            sheet.Cells["B26:D26"].Value = "TOTAL NO. OF AFB RECEIVED:";

            sheet.Cells["I26"].Formula = "SUM(I5:I15,I24)";
            sheet.Cells["N26"].Formula = "SUM(N5:N15,N24)";
            sheet.Cells["S26"].Formula = "SUM(S5:S15,S24)";
            sheet.Cells["X26"].Formula = "SUM(X5:X15,X24)";

            sheet.Cells["Y26"].Formula = "SUM(Y16,Y24)";
            sheet.Cells["Y26"].Style.Font.Color.SetColor(Color.Purple);
            sheet.Cells["B26:Y26"].Style.Border.BorderAround(ExcelBorderStyle.Thick, Color.Black);

            #endregion

            #region GS Category
            row = 32;

            specimens = monthlyReportEntity.TestTallies[1][1].CategoryTallies.Select(x => x.SpecimenType).ToList();

            foreach (var specimen in specimens)
            {
                string counterLabel = (row - 31).ToString() + ". ";
                sheet.Cells[$"B{row}:D{row}"].Merge = true;
                if (row >= 47) break;

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
            sheet.Cells[$"B47:D47"].Merge = true;
            sheet.Cells[$"B47:D47"].Value = "TOTAL NO OF. SPECIMEN RECEIVED:";
            sheet.Cells["B47:Y47"].Style.Border.BorderAround(ExcelBorderStyle.Thick, Color.Black);

            sheet.Cells["I47"].Formula = "=SUM(I32:I45)";
            sheet.Cells["N47"].Formula = "=SUM(N32:N45)";
            sheet.Cells["S47"].Formula = "=SUM(S32:S45)";
            sheet.Cells["X47"].Formula = "=SUM(X32:X45)";
            sheet.Cells["Y47"].Formula = "=SUM(Y32:Y45)";

            sheet.Cells["I47"].Style.Font.Color.SetColor(Color.Red);
            sheet.Cells["N47"].Style.Font.Color.SetColor(Color.Red);
            sheet.Cells["S47"].Style.Font.Color.SetColor(Color.Red);
            sheet.Cells["X47"].Style.Font.Color.SetColor(Color.Red);
            sheet.Cells["Y47"].Style.Font.Color.SetColor(Color.Purple);

            #endregion

            #region C/S Category
            row = 53;

            specimens = monthlyReportEntity.TestTallies[1][2].CategoryTallies.Select(x => x.SpecimenType).ToList();
            int ctrLabel = 1, incrementFactor = 0;
            foreach (var specimen in specimens)
            {
                string specimenLabel = $"{ctrLabel}. {specimen}";
                sheet.Cells[$"B{row}:D{row}"].Merge = true;

                if (row - 52 == 1)
                {
                    specimenLabel = $"{ctrLabel}. CSF,AF,PF";
                    sheet.Cells[$"B{row}:D{row}"].Value = specimenLabel;
                    row++;
                    sheet.Cells[$"B{row}:D{row}"].Merge = true;
                }

                if (row < 63 && row > 53)
                {
                    specimenLabel = specimen;
                    sheet.Cells[$"B{row}:D{row}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    if (row == 61)
                    {
                        specimenLabel = "TOTAL TRANSUDATES";

                        sheet.Cells[$"B{row}:D{row}"].Value = specimenLabel;
                        sheet.Cells[$"B{row}:D{row}"].Style.Font.Bold = false;
                        sheet.Cells[$"B{row}:D{row}"].Style.Font.Italic = true;

                        sheet.Cells[$"E{row}"].Formula = "=SUM(E54:E60)";
                        sheet.Cells[$"F{row}"].Formula = "=SUM(F54:F60)";
                        sheet.Cells[$"G{row}"].Formula = "=SUM(G54:G60)";
                        sheet.Cells[$"H{row}"].Formula = "=SUM(H54:H60)";
                        sheet.Cells[$"I{row}"].Formula = "=SUM(I54:I60)";

                        sheet.Cells[$"J{row}"].Formula = "=SUM(J54:J60)";
                        sheet.Cells[$"K{row}"].Formula = "=SUM(K54:K60)";
                        sheet.Cells[$"L{row}"].Formula = "=SUM(L54:L60)";
                        sheet.Cells[$"M{row}"].Formula = "=SUM(M54:M60)";
                        sheet.Cells[$"N{row}"].Formula = "=SUM(N54:N60)";

                        sheet.Cells[$"O{row}"].Formula = "=SUM(O54:O60)";
                        sheet.Cells[$"P{row}"].Formula = "=SUM(P54:P60)";
                        sheet.Cells[$"Q{row}"].Formula = "=SUM(Q54:Q60)";
                        sheet.Cells[$"R{row}"].Formula = "=SUM(R54:R60)";
                        sheet.Cells[$"S{row}"].Formula = "=SUM(S54:S60)";

                        sheet.Cells[$"T{row}"].Formula = "=SUM(T54:T60)";
                        sheet.Cells[$"U{row}"].Formula = "=SUM(U54:U60)";
                        sheet.Cells[$"V{row}"].Formula = "=SUM(V54:V60)";
                        sheet.Cells[$"W{row}"].Formula = "=SUM(W54:W60)";
                        sheet.Cells[$"X{row}"].Formula = "=SUM(X54:X60)";

                        sheet.Cells[$"Y{row}"].Formula = "=SUM(Y54:Y60)";
                        ctrLabel++;
                        row++;
                        sheet.Cells[$"B{row}:D{row}"].Merge = true;
                        row++;
                        specimenLabel = $"{ctrLabel}. {specimen}";
                        sheet.Cells[$"B{row}:D{row}"].Merge = true;
                        incrementFactor = 1;
                    }
                }

                if (row == 66)
                {
                    incrementFactor = 0;
                    specimenLabel = "5. WOUND/ABSCESS";
                    sheet.Cells[$"B{row}:D{row}"].Value = specimenLabel;
                    row++;
                    specimenLabel = specimen;
                }


                if (row > 66 && row <= 70)
                {
                    specimenLabel = specimen;
                    sheet.Cells[$"B{row}:D{row}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    if (row == 70)
                    {
                        specimenLabel = "TOTAL EXUDATES";

                        sheet.Cells[$"B{row}:D{row}"].Value = specimenLabel;
                        sheet.Cells[$"B{row}:D{row}"].Style.Font.Bold = false;
                        sheet.Cells[$"B{row}:D{row}"].Style.Font.Italic = true;

                        sheet.Cells[$"E{row}"].Formula = "=SUM(E67:E69)";
                        sheet.Cells[$"F{row}"].Formula = "=SUM(F67:F69)";
                        sheet.Cells[$"G{row}"].Formula = "=SUM(G67:G69)";
                        sheet.Cells[$"H{row}"].Formula = "=SUM(H67:H69)";
                        sheet.Cells[$"I{row}"].Formula = "=SUM(I67:I69)";

                        sheet.Cells[$"J{row}"].Formula = "=SUM(J67:J69)";
                        sheet.Cells[$"K{row}"].Formula = "=SUM(K67:K69)";
                        sheet.Cells[$"L{row}"].Formula = "=SUM(L67:L69)";
                        sheet.Cells[$"M{row}"].Formula = "=SUM(M67:M69)";
                        sheet.Cells[$"N{row}"].Formula = "=SUM(N67:N69)";

                        sheet.Cells[$"O{row}"].Formula = "=SUM(O67:O69)";
                        sheet.Cells[$"P{row}"].Formula = "=SUM(P67:P69)";
                        sheet.Cells[$"Q{row}"].Formula = "=SUM(Q67:Q69)";
                        sheet.Cells[$"R{row}"].Formula = "=SUM(R67:R69)";
                        sheet.Cells[$"S{row}"].Formula = "=SUM(S67:S69)";

                        sheet.Cells[$"T{row}"].Formula = "=SUM(T67:T69)";
                        sheet.Cells[$"U{row}"].Formula = "=SUM(U67:U69)";
                        sheet.Cells[$"V{row}"].Formula = "=SUM(V67:V69)";
                        sheet.Cells[$"W{row}"].Formula = "=SUM(W67:W69)";
                        sheet.Cells[$"X{row}"].Formula = "=SUM(X67:X69)";

                        sheet.Cells[$"Y{row}"].Formula = "=SUM(Y67:Y69)";
                        incrementFactor = 1;
                        ctrLabel += incrementFactor;
                        row++;
                        sheet.Cells[$"B{row}:D{row}"].Merge = true;
                    }
                }

                if (row == 71)
                {
                    incrementFactor = 0;
                    specimenLabel = "6. RESPIRATORY DISCHARGE";
                    sheet.Cells[$"B{row}:D{row}"].Value = specimenLabel;
                    row++;
                    specimenLabel = specimen;
                }

                if (row > 71 && row <= 74)
                {
                    specimenLabel = specimen;
                    sheet.Cells[$"B{row}:D{row}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    if (row == 74)
                    {
                        specimenLabel = "TOTAL RESPIRATORY";

                        sheet.Cells[$"B{row}:D{row}"].Value = specimenLabel;
                        sheet.Cells[$"B{row}:D{row}"].Style.Font.Bold = false;
                        sheet.Cells[$"B{row}:D{row}"].Style.Font.Italic = true;

                        sheet.Cells[$"E{row}"].Formula = "=SUM(E72:E73)";
                        sheet.Cells[$"F{row}"].Formula = "=SUM(F72:F73)";
                        sheet.Cells[$"G{row}"].Formula = "=SUM(G72:G73)";
                        sheet.Cells[$"H{row}"].Formula = "=SUM(H72:H73)";
                        sheet.Cells[$"I{row}"].Formula = "=SUM(I72:I73)";

                        sheet.Cells[$"J{row}"].Formula = "=SUM(J72:J73)";
                        sheet.Cells[$"K{row}"].Formula = "=SUM(K72:K73)";
                        sheet.Cells[$"L{row}"].Formula = "=SUM(L72:L73)";
                        sheet.Cells[$"M{row}"].Formula = "=SUM(M72:M73)";
                        sheet.Cells[$"N{row}"].Formula = "=SUM(N72:N73)";

                        sheet.Cells[$"O{row}"].Formula = "=SUM(O72:O73)";
                        sheet.Cells[$"P{row}"].Formula = "=SUM(P72:P73)";
                        sheet.Cells[$"Q{row}"].Formula = "=SUM(Q72:Q73)";
                        sheet.Cells[$"R{row}"].Formula = "=SUM(R72:R73)";
                        sheet.Cells[$"S{row}"].Formula = "=SUM(S72:S73)";

                        sheet.Cells[$"T{row}"].Formula = "=SUM(T72:T73)";
                        sheet.Cells[$"U{row}"].Formula = "=SUM(U72:U73)";
                        sheet.Cells[$"V{row}"].Formula = "=SUM(V72:V73)";
                        sheet.Cells[$"W{row}"].Formula = "=SUM(W72:W73)";
                        sheet.Cells[$"X{row}"].Formula = "=SUM(X72:X73)";

                        sheet.Cells[$"Y{row}"].Formula = "=SUM(Y72:Y73)";
                        incrementFactor = 1;
                        ctrLabel += incrementFactor;
                        row++;
                        sheet.Cells[$"B{row}:D{row}"].Merge = true;
                        row++;
                        specimenLabel = specimen;
                    }
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
                ctrLabel += incrementFactor;
            }

            sheet.Cells[$"B{row}:D{row}"].Merge = true;
            row++;

            sheet.Cells["B78:Y78"].Style.Border.BorderAround(ExcelBorderStyle.Thick, Color.Black);
            sheet.Cells[$"B78:D78"].Merge = true;
            sheet.Cells[$"B78:D78"].Value = "TOTAL NO OF. CULTURE & SENSI RECEIVED:";
            sheet.Cells["B78:Y78"].Style.Border.BorderAround(ExcelBorderStyle.Thick, Color.Black);

            sheet.Cells["I78"].Formula = "=SUM(I61,I63:I65,I70,I74,I76)";
            sheet.Cells["N78"].Formula = "=SUM(N61,N63:N65,N70,N74,N76)";
            sheet.Cells["S78"].Formula = "=SUM(S61,S63:S65,S70,S74,S76)";
            sheet.Cells["X78"].Formula = "=SUM(X61,X63:X65,X70,X74,X76)";
            sheet.Cells["Y78"].Formula = "=SUM(Y61,Y63:Y65,Y70,Y74,Y76)";

            sheet.Cells["I78"].Style.Font.Color.SetColor(Color.Red);
            sheet.Cells["N78"].Style.Font.Color.SetColor(Color.Red);
            sheet.Cells["S78"].Style.Font.Color.SetColor(Color.Red);
            sheet.Cells["X78"].Style.Font.Color.SetColor(Color.Red);
            sheet.Cells["Y78"].Style.Font.Color.SetColor(Color.Purple);


            #endregion

            #region Other Categories
            row = 84;

            specimens = monthlyReportEntity.TestTallies[1].Skip(3).Select(x => x.CategoryName).ToList();
            foreach (var specimen in specimens)
            {
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

            sheet.Cells[$"B88:D88"].Merge = true;
            row++;
            sheet.Cells[$"B89:D89"].Merge = true;
            sheet.Cells[$"B89:D89"].Value = "TOTAL NO OF. SPECIMEN RECEIVED:";
            sheet.Cells["B89:Y89"].Style.Border.BorderAround(ExcelBorderStyle.Thick, Color.Black);

            sheet.Cells["I89"].Formula = "=SUM(I84:I87)";
            sheet.Cells["N89"].Formula = "=SUM(N84:N87)";
            sheet.Cells["S89"].Formula = "=SUM(S84:S87)";
            sheet.Cells["X89"].Formula = "=SUM(X84:X87)";
            sheet.Cells["Y89"].Formula = "=SUM(Y84:Y87)";

            sheet.Cells["I89"].Style.Font.Color.SetColor(Color.Red);
            sheet.Cells["N89"].Style.Font.Color.SetColor(Color.Red);
            sheet.Cells["S89"].Style.Font.Color.SetColor(Color.Red);
            sheet.Cells["X89"].Style.Font.Color.SetColor(Color.Red);
            sheet.Cells["Y89"].Style.Font.Color.SetColor(Color.Purple);


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
            sheet.Cells["A5:C5"].Value = "Total Specimen Received";
            sheet.Cells["A5:C5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            sheet.Cells["A22:C23,A39:C40,A50:C51"].Style.Font.Bold = true;
            sheet.Cells["A22:C23,A39:C40,A50:C51"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


            sheet.Cells["D4:H5,D8:H21,D23:H38,D40:H49,D51:H57"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            sheet.Cells["D4:H5,D8:H21,D23:H38,D40:H49,D51:H57"].Style.Border.Top.Color.SetColor(Color.Black);

            sheet.Cells["D4:H5,D8:H21,D23:H38,D40:H49,D51:H57"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            sheet.Cells["D4:H5,D8:H21,D23:H38,D40:H49,D51:H57"].Style.Border.Left.Color.SetColor(Color.Black);

            sheet.Cells["D4:H5,D8:H21,D23:H38,D40:H49,D51:H57"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            sheet.Cells["D4:H5,D8:H21,D23:H38,D40:H49,D51:H57"].Style.Border.Right.Color.SetColor(Color.Black);

            sheet.Cells["D4:H5,D8:H21,D23:H38,D40:H49,D51:H57"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            sheet.Cells["D4:H5,D8:H21,D23:H38,D40:H49,D51:H57"].Style.Border.Bottom.Color.SetColor(Color.Black);

            sheet.Cells["D3,D7,D23,D40,D51"].Value = "IP";
            sheet.Cells["E3,E7,E23,E40,E51"].Value = "ER";
            sheet.Cells["F3,F7,F23,F40,F51"].Value = "MAGS";
            sheet.Cells["G3,G7,G23,G40,G51"].Value = "OPD";
            sheet.Cells["H3,H7,H23,H40,H51"].Value = "TOTAL";


            sheet.Cells["D3:H58"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Cells["H3:H5,H7:H21,H23:H38,H40:H49,H51:H57"].Style.Font.Bold = true;

            sheet.Cells["D3:H3,D7:H7,D23:H23,D40:H40,D51:H51"].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);


            sheet.Cells["A8:C8"].Style.Font.Bold = true;
            sheet.Cells["A8:C8"].Merge = true;
            sheet.Cells["A8:C8"].Value = "AFB";
            sheet.Cells["A8:C8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            int row = 9;
            int refRow = 5;

            #region AFB
            var aggregateSpecimen = monthlyReportEntity.TestTallies[1][0].CategoryTallies.Select(x => x.SpecimenType).ToList();
            for (int i = 0; i < 12; i++)
            {
                sheet.Cells[$"A{row}:C{row}"].Merge = true;

                if (i == 11)
                {

                    refRow = 24;
                    sheet.Cells[$"A{row}:C{row}"].Value = $"{i + 1}. OTHERS";

                    sheet.Cells[$"D{row}"].Formula = $"=SUM(WEEKLY!E{refRow}, WEEKLY!J{refRow}, WEEKLY!O{refRow}, WEEKLY!T{refRow})";
                    sheet.Cells[$"E{row}"].Formula = $"=SUM(WEEKLY!F{refRow}, WEEKLY!K{refRow}, WEEKLY!P{refRow}, WEEKLY!U{refRow})";
                    sheet.Cells[$"F{row}"].Formula = $"=SUM(WEEKLY!G{refRow}, WEEKLY!L{refRow}, WEEKLY!Q{refRow}, WEEKLY!V{refRow})";
                    sheet.Cells[$"G{row}"].Formula = $"=SUM(WEEKLY!H{refRow}, WEEKLY!M{refRow}, WEEKLY!R{refRow}, WEEKLY!W{refRow})";
                    sheet.Cells[$"H{row}"].Formula = $"=SUM(D{row}:G{row})";

                    row++;

                    sheet.Cells[$"D{row}"].Formula = $"=SUM(D9:D20)";
                    sheet.Cells[$"E{row}"].Formula = $"=SUM(E9:E20)";
                    sheet.Cells[$"F{row}"].Formula = $"=SUM(F9:F20)";
                    sheet.Cells[$"G{row}"].Formula = $"=SUM(G9:G20)";
                    sheet.Cells[$"H{row}"].Formula = $"=SUM(H9:H20)";
                    sheet.Cells[$"H{row}"].Style.Font.Color.SetColor(Color.Red);

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
            row = 24;
            refRow = 32;

            sheet.Cells[$"A22:C23"].Merge = true;
            sheet.Cells[$"A22:C22"].Value = "GRAM STAINING";
            sheet.Cells[$"A23:C23"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            aggregateSpecimen = monthlyReportEntity.TestTallies[1][1].CategoryTallies.Select(x => x.SpecimenType).ToList();
            for (int i = 0; i < 14; i++)
            {
                sheet.Cells[$"A{row}:C{row}"].Merge = true;


                sheet.Cells[$"A{row}:C{row}"].Value = $"{i + 1}. {aggregateSpecimen[i]}";

                sheet.Cells[$"D{row}"].Formula = $"=SUM(WEEKLY!E{refRow}, WEEKLY!J{refRow}, WEEKLY!O{refRow}, WEEKLY!T{refRow})";
                sheet.Cells[$"E{row}"].Formula = $"=SUM(WEEKLY!F{refRow}, WEEKLY!K{refRow}, WEEKLY!P{refRow}, WEEKLY!U{refRow})";
                sheet.Cells[$"F{row}"].Formula = $"=SUM(WEEKLY!G{refRow}, WEEKLY!L{refRow}, WEEKLY!Q{refRow}, WEEKLY!V{refRow})";
                sheet.Cells[$"G{row}"].Formula = $"=SUM(WEEKLY!H{refRow}, WEEKLY!M{refRow}, WEEKLY!R{refRow}, WEEKLY!W{refRow})";
                sheet.Cells[$"H{row}"].Formula = $"=SUM(D{row}:G{row})";

                if (i == 13)
                {

                    row++;
                    sheet.Cells[$"D{row}"].Formula = $"=SUM(D24:D37)";
                    sheet.Cells[$"E{row}"].Formula = $"=SUM(E24:E37)";
                    sheet.Cells[$"F{row}"].Formula = $"=SUM(F24:F37)";
                    sheet.Cells[$"G{row}"].Formula = $"=SUM(G24:G37)";
                    sheet.Cells[$"H{row}"].Formula = $"=SUM(H24:H37)";
                    sheet.Cells[$"H{row}"].Style.Font.Color.SetColor(Color.Red);

                }

                refRow++;
                row++;
            }

            #endregion

            #region C/S

            row = 41;
            refRow = 32;

            sheet.Cells[$"A39:C40"].Merge = true;
            sheet.Cells[$"A39:C39"].Value = "CULTURE AND SENSITIVITY";
            sheet.Cells[$"A23:C23"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            for (int i = 0; i < 7; i++)
            {
                sheet.Cells[$"A{row}:C{row}"].Merge = true;
                string specimenLabel = string.Empty;

                switch (i + 1)
                {
                    case 1:
                        refRow = 61;
                        specimenLabel = "CSF,AF,DF";
                        break;
                    case 2:
                        refRow = 63;
                        specimenLabel = "RECTAL SWAB/STOOL";
                        break;
                    case 3:
                        refRow = 64;
                        specimenLabel = "URINE";
                        break;
                    case 4:
                        refRow = 65;
                        specimenLabel = "BLOOD";
                        break;
                    case 5:
                        refRow = 70;
                        specimenLabel = "WOUND/ABSCESS";
                        break;
                    case 6:
                        refRow = 74;
                        specimenLabel = "RESPIRATORY DISCHARGE";
                        break;
                    case 7:

                        row++;
                        sheet.Cells[$"A{row}:C{row}"].Merge = true;
                        refRow = 76;
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
            sheet.Cells[$"D{row}"].Formula = $"=SUM(D41:D48)";
            sheet.Cells[$"E{row}"].Formula = $"=SUM(E41:E48)";
            sheet.Cells[$"F{row}"].Formula = $"=SUM(F41:F48)";
            sheet.Cells[$"G{row}"].Formula = $"=SUM(G41:G48)";
            sheet.Cells[$"H{row}"].Formula = $"=SUM(H41:H48)";
            sheet.Cells[$"H{row}"].Style.Font.Color.SetColor(Color.Red);

            sheet.Cells[$"H{row}"].Style.Font.Color.SetColor(Color.Red);

            #endregion

            #region Other Categories 


            row = 52;
            refRow = 84;

            sheet.Cells[$"A50:C51"].Merge = true;
            sheet.Cells[$"A50:C51"].Value = "GRAM STAINING";
            sheet.Cells[$"A50:C51"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

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
            row++;
            sheet.Cells[$"D{row}"].Formula = $"=SUM(D52:D55)";
            sheet.Cells[$"E{row}"].Formula = $"=SUM(E52:E55)";
            sheet.Cells[$"F{row}"].Formula = $"=SUM(F52:F55)";
            sheet.Cells[$"G{row}"].Formula = $"=SUM(G52:G55)";
            sheet.Cells[$"H{row}"].Formula = $"=SUM(H52:H55)";
            sheet.Cells[$"H{row}"].Style.Font.Color.SetColor(Color.Red);


            #endregion  
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

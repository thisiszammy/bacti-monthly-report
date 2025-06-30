using MonthlyReportEntrySoftware.Entities.BaseEntities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MonthlyReportEntrySoftware.Entities
{
    public class MonthlyReportEntity
    {
        public int MonthCode { get; set; }
        public int Year { get; set; }
        public IDictionary<int, List<TestCategory>> TestTallies { get; set; }
        string fileName;
        public string FileName
        {
            get
            {
                return fileName;
            }
            set
            {
                if (value.Contains(".xlsx")) fileName = value.Replace(".xlsx", "");
                else if (value.Contains(".xls")) fileName = value.Replace(".xls", "");
                else fileName = value;
            }
        }
        public string FilePath { get; set; }
        public string FullPath { get => FilePath + @"\" + FileName + ".xlsx"; }

        public MonthlyReportEntity()
        {
            GenerateTestCategories();
        }

        public MonthlyReportEntity(int MonthCode, int Year)
        {
            this.MonthCode = MonthCode;
            this.Year = Year;
            GenerateTestCategories();
        }

        private void GenerateTestCategories()
        {
            TestTallies = new Dictionary<int, List<TestCategory>>();

            TestTallies.Add(1, new List<TestCategory>
            {
                new TestCategory(TestEnum.AFB),
                new TestCategory(TestEnum.GS),
                new TestCategory(TestEnum.CS),
                new TestCategory(TestEnum.HM),
                new TestCategory(TestEnum.GE),
                new TestCategory(TestEnum.KOH),
                new TestCategory(TestEnum.INDIA),
                new TestCategory(TestEnum.XDR)
            });

            TestTallies.Add(2, new List<TestCategory>
            {
                new TestCategory(TestEnum.AFB),
                new TestCategory(TestEnum.GS),
                new TestCategory(TestEnum.CS),
                new TestCategory(TestEnum.HM),
                new TestCategory(TestEnum.GE),
                new TestCategory(TestEnum.KOH),
                new TestCategory(TestEnum.INDIA),
                new TestCategory(TestEnum.XDR)
            });

            TestTallies.Add(3, new List<TestCategory>
            {
                new TestCategory(TestEnum.AFB),
                new TestCategory(TestEnum.GS),
                new TestCategory(TestEnum.CS),
                new TestCategory(TestEnum.HM),
                new TestCategory(TestEnum.GE),
                new TestCategory(TestEnum.KOH),
                new TestCategory(TestEnum.INDIA),
                new TestCategory(TestEnum.XDR)
            });

            TestTallies.Add(4, new List<TestCategory>
            {
                new TestCategory(TestEnum.AFB),
                new TestCategory(TestEnum.GS),
                new TestCategory(TestEnum.CS),
                new TestCategory(TestEnum.HM),
                new TestCategory(TestEnum.GE),
                new TestCategory(TestEnum.KOH),
                new TestCategory(TestEnum.INDIA),
                new TestCategory(TestEnum.XDR)
            });

        }
    }
}

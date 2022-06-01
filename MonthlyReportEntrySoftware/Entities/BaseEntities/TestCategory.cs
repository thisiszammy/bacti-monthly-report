using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MonthlyReportEntrySoftware.Entities.BaseEntities
{
    public class TestCategory
    {
        public string CategoryName { get; set; }
        public List<CategoryTally> CategoryTallies { get; set; }
        public int CategoryTallyTotal { get => CategoryTallies == null ? 0 : CategoryTallies.Sum(x => x.Total); }

        public TestCategory(TestEnum testType)
        {
            switch (testType)
            {
                case TestEnum.AFB:
                    this.CategoryName = "AFB";
                    CategoryTallies = new List<CategoryTally>
                    {
                        new CategoryTally("ASCITIC FLUID"),
                        new CategoryTally("CSF"),
                        new CategoryTally("PERITONEAL FLUID"),
                        new CategoryTally("PLEURAL FLUID"),
                        new CategoryTally("PERICARDIAL FLUID"),
                        new CategoryTally("PROSTATIC DISCHARGES"),
                        new CategoryTally("SKIN LESIONS"),
                        new CategoryTally("SLIT & SCRAPE"),
                        new CategoryTally("SPUTUM"),
                        new CategoryTally("URINE SEDIMENTS"),
                        new CategoryTally("STOOL"),
                        new CategoryTally("WOUND"),
                        new CategoryTally("UVC TIP/NG TTIP"),
                        new CategoryTally("ABSCESS"),
                        new CategoryTally("TISSUE"),
                        new CategoryTally("TA"),
                        new CategoryTally("ASPIRATE & FLUIDS")
                    };
                    break;
                case TestEnum.GS:
                    this.CategoryName = "GRAM STAIN";
                    CategoryTallies = new List<CategoryTally>
                    {
                        new CategoryTally("ASCITIC FLUID"),
                        new CategoryTally("CSF"),
                        new CategoryTally("EAR DISCHARGE"),
                        new CategoryTally("PLEURAL/PERITONEAL FLUID"),
                        new CategoryTally("PERICARDIAL FLUID"),
                        new CategoryTally("SKIN LESIONS"),
                        new CategoryTally("SPUTUM"),
                        new CategoryTally("THROAT SWAB"),
                        new CategoryTally("URETHRAL SMEAR/PENILE SMEAR"),
                        new CategoryTally("URINE"),
                        new CategoryTally("VAGINAL SMEAR"),
                        new CategoryTally("WOUND DISCHARGE/ABSCESS"),
                        new CategoryTally("ETA"),
                        new CategoryTally("OTHERS")
                    };
                    break;
                case TestEnum.CS:
                    this.CategoryName = "CULTURE & SENSI";
                    CategoryTallies = new List<CategoryTally>
                    {
                        new CategoryTally("CSF"),
                        new CategoryTally("AF"),
                        new CategoryTally("PF"),
                        new CategoryTally("PLU"),
                        new CategoryTally("PRECARDIAL"),
                        new CategoryTally("GASTRIC FLUID"),
                        new CategoryTally("OTHERS (CSF,AF,PF)"),
                        new CategoryTally("RECTAL SWAB/STOOL"),
                        new CategoryTally("URINE"),
                        new CategoryTally("BLOOD"),
                        new CategoryTally("WOUND"),
                        new CategoryTally("ABSCESS"),
                        new CategoryTally("OTHER DISCHARGES"),
                        new CategoryTally("ETA"),
                        new CategoryTally("SPUTUM"),
                        new CategoryTally("ENVIRONMENTAL CULTURE")
                    };
                    break;
                case TestEnum.HM:
                    this.CategoryName = "HUMAN MILK";
                    CategoryTallies = new List<CategoryTally>
                    {
                        new CategoryTally("HUMAN MILK")
                    };
                    break;
                case TestEnum.GE:
                    this.CategoryName = "GENE EXPERT";
                    CategoryTallies = new List<CategoryTally>
                    {
                        new CategoryTally("GENE EXPERT")
                    };
                    break;
                case TestEnum.KOH:
                    this.CategoryName = "KOH";
                    CategoryTallies = new List<CategoryTally>
                    {
                        new CategoryTally("KOH")
                    };
                    break;
                case TestEnum.INDIA:
                    this.CategoryName = "INDIA INK";
                    CategoryTallies = new List<CategoryTally>
                    {
                        new CategoryTally("INDIA INK PREP")
                    };
                    break;
            }
        }
    }
}

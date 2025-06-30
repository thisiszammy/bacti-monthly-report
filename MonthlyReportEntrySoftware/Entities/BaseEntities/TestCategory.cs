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
                        new CategoryTally("BILE"),
                        new CategoryTally("OTHERS:", true),
                        new CategoryTally("WOUND"),
                        new CategoryTally("UVC TIP/NG TTIP"),
                        new CategoryTally("TISSUE"),
                        new CategoryTally("ENDOTRACHEAL ASPIRATE"),
                        new CategoryTally("EXUDATES"),
                        new CategoryTally("THROAT SWAB")
                    };
                    break;
                case TestEnum.GS:
                    this.CategoryName = "GRAM STAIN";
                    CategoryTallies = new List<CategoryTally>
                    {
                        new CategoryTally("ASCITIC FLUID"),
                        new CategoryTally("CSF"),
                        new CategoryTally("EAR DISCHARGE"),
                        new CategoryTally("PLEURAL FLUID"),
                        new CategoryTally("PERITONEAL FLUID"),
                        new CategoryTally("BILE"),
                        new CategoryTally("SYNOVIAL"),
                        new CategoryTally("PERICARDIAL FLUID"),
                        new CategoryTally("PANCREATIC FLUID"),
                        new CategoryTally("SPUTUM"),
                        new CategoryTally("THROAT SWAB"),
                        new CategoryTally("URETHRAL SMEAR/PENILE SMEAR"),
                        new CategoryTally("URINE"),
                        new CategoryTally("VAGINAL DISCHARGE/RECTOVAGINAL"),
                        new CategoryTally("TISSUE/MASS"),
                        new CategoryTally("WOUND DISCHARGE/ABSCESS"),
                        new CategoryTally("ETA"),
                        new CategoryTally("BLOOD"),
                        new CategoryTally("CATHETERS"),
                        new CategoryTally("SEMEN")
                    };
                    break;
                case TestEnum.CS:
                    this.CategoryName = "CULTURE & SENSI";
                    CategoryTallies = new List<CategoryTally>
                    {
                        new CategoryTally("CSF,AF,PF", true),
                        new CategoryTally("CSF"),
                        new CategoryTally("ASCITIC/GASTRIC"),
                        new CategoryTally("PERITONEAL"),
                        new CategoryTally("PLEURAL"),
                        new CategoryTally("PRECARDIAL"),
                        new CategoryTally("PANCREATIC FLUID"),
                        new CategoryTally("BILE"),
                        new CategoryTally("SYNOVIAL"),
                        new CategoryTally("SEMEN"),
                        new CategoryTally("OTHERS (CSF,AF,PF)"),
                        new CategoryTally("RECTAL SWAB/STOOL"),
                        new CategoryTally("URINE"),
                        new CategoryTally("BLOOD"),
                        new CategoryTally("WOUND/ABSCESS", true),
                        new CategoryTally("WOUND/ABSCESS"),
                        new CategoryTally("VAGINAL DISCHARGE/RECTOVAGINAL SWAB"),
                        new CategoryTally("TISSUE/MASS"),
                        new CategoryTally("CATHETERS"),
                        new CategoryTally("EXUDATES"),
                        new CategoryTally("PENILE/URETHRAL DISCHARGE"),
                        new CategoryTally("OTHER DISCHARGES"),
                        new CategoryTally("RESPIRATORY DISCHARGE", true),
                        new CategoryTally("ETA"),
                        new CategoryTally("SPUTUM"),
                        new CategoryTally("THROAT SWAB"),
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
                    this.CategoryName = "INDIA";
                    CategoryTallies = new List<CategoryTally>
                    {
                        new CategoryTally("INDIA")
                    };
                    break;
                case TestEnum.XDR:
                    this.CategoryName = "XDR";
                    CategoryTallies = new List<CategoryTally>
                    {
                        new CategoryTally("XDR")
                    };
                    break;
            }
        }
    }
}

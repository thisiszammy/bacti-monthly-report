using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MonthlyReportEntrySoftware.Entities.BaseEntities
{
    public class CategoryTally
    {
        public string SpecimenType { get; set; }
        public int IP { get; set; }
        public int ER { get; set; }
        public int MAGS { get; set; }
        public int OPD { get; set; }
        public int Total { get => IP + ER + MAGS + OPD; }

        public bool IsHeader { get; set; }

        public CategoryTally(string SpecimenType, bool isHeader = false)
        {
            this.SpecimenType = SpecimenType;
            IsHeader = isHeader;
            IP = 0;
            ER = 0;
            MAGS = 0;
            OPD = 0;
        }
    }
}

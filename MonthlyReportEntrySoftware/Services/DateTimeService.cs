using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MonthlyReportEntrySoftware.Services
{
    public static class DateTimeService
    {
        private static string[] Months =
        {
            "JANUARY",
            "FEBRUARY",
            "MARCH",
            "APRIL",
            "MAY",
            "JUNE",
            "JULY",
            "AUGUST",
            "SEPTEMBER",
            "OCTOBER",
            "NOVEMBER",
            "DECEMBER"
        };

        public static string GetMonthFromCode(int code)
        {
            if (code < 1 || code > 12) throw new IndexOutOfRangeException("Invalid Month Code!");
            return Months[code-1];
        }

        public static int GetMonthCodeFromString(string month)
        {
            for(int i = 1; i <= Months.Count(); i++)
            {
                if (Months[i-1] == month) return i;
            }
            throw new IndexOutOfRangeException();
        }

        public static string GetDefaultDateTimeFormat()
        {
            return CultureInfo.CurrentUICulture.DateTimeFormat.ShortDatePattern;
        }

        public static List<string> GetMonths()
        {
            return new List<string>(Months);
        }

        public static List<int> GenerateYearSpan(int span)
        {
            int middle = span / 2 + ((span % 2 == 0) ? 0 : 1);
            List<int> yearList = new List<int>();

            for(int i = 1; i < middle; i++)
            {
                yearList.Add(DateTime.Now.Year - (middle - i));
            }

            yearList.Add(DateTime.Now.Year);

            for(int i = middle+1; i <= span; i++)
            {
                yearList.Add(DateTime.Now.Year + i - middle);
            }

            return yearList;
        }
    }
}

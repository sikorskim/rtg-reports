using System;
using System.Collections.Generic;

namespace computerman_rtg_reports
{
    public class UnitsReport
    {
        public int Id { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public string Unit { get; set; }

        public List<UnitsReportItem> Items { get; set; }

    }
}
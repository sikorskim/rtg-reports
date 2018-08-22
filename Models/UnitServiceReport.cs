using System;

namespace computerman_rtg_reports
{
    public class UnitServiceReportItem
    {
        public int Id { get; set; }
        public string Unit { get; set; }
        public string Code { get; set; }
        public int Count { get; set; }
        public int Photos { get; set; }
        public decimal Value { get; set; }
    }
}
using System;

namespace computerman_rtg_reports
{
    public class ServiceReportItem
    {
        public int Id { get; set; }
        public string Code { get; set; }
        public string Name { get; set; }
        public int Count { get; set; }
        public int Photos { get; set; }
        public decimal Value { get; set; }
    }
}
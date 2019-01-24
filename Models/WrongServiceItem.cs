using System;

namespace computerman_rtg_reports
{
    public class WrongServiceItem
    {
        public int Id { get; set; }
        public string Unit { get; set; }
        public string Code { get; set; }
        public string PatientName { get; set; }
        public string PatientPesel { get; set; }
        public DateTime Date { get; set; }
    }
}
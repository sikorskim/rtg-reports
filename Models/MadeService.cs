using System;

namespace computerman_rtg_reports
{
    public class MadeService
    {
        public int Id { get; set; }
        public DateTime Date { get; set; }
        public string PatientName { get; set; }
        public string PatientPesel { get; set; }
        public Service Service { get; set; }
    }
}
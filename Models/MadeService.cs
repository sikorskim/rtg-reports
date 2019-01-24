using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace computerman_rtg_reports
{
    public class MadeService
    {
        public int Id { get; set; }
        public DateTime Date { get; set; }
        public string PatientName { get; set; }
        public string PatientPesel { get; set; }
        public string ServiceCode { get; set; }
        public string Unit { get; set; }        

        public string getServiceCode(string input)
        {
            return input.Substring(0, input.IndexOf(' '));  
        }

        public string getUnit(string input)
        {
            if(input==" ")
            {
                input="Skierowanie zewnętrzne";
            }

            return input;
        }

        public static List<MadeService> getMadeServices(string filename)
        {
            FileInfo file = new FileInfo(filename);
            List<MadeService> madeServList = new List<MadeService>();

            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();
                int rowCount = worksheet.Dimension.Rows;
                int startRow = 6;
                string dtFormat = "dd-MM-yyyy";

                for (int row = startRow; row <= rowCount; row++)
                {
                    MadeService ms = new MadeService();
                    ms.Id = Int32.Parse(worksheet.Cells[row, 1].Value.ToString());
                    ms.Date = DateTime.ParseExact(worksheet.Cells[row, 2].Value.ToString(), dtFormat, CultureInfo.InvariantCulture);
                    ms.PatientName = worksheet.Cells[row, 6].Value.ToString();
                    ms.PatientPesel = worksheet.Cells[row, 8].Value.ToString();
                    ms.ServiceCode = ms.getServiceCode(worksheet.Cells[row, 11].Value.ToString());
                    ms.Unit = ms.getUnit(worksheet.Cells[row, 10].Value.ToString());

                    madeServList.Add(ms);
                }
            }
            return madeServList;
        }
    }
}
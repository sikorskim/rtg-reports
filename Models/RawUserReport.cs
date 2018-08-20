using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace computerman_rtg_reports
{
    public class RawUserReport
    {
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public string Unit { get; set; }
        public List<MadeService> MadeServicesList { get; set; }
        ExcelWorksheet excelWorksheet;
        string reportMetadata;

        public RawUserReport (string filename)
        {
            excelWorksheet = getExcelWorksheet(filename);
            reportMetadata = getMetadata ();
            StartDate = getStartDate (reportMetadata);
            EndDate = getEndDate (reportMetadata);
            Unit = getUnit (reportMetadata);            
            MadeServicesList = getMadeServices (filename);
        }

        ExcelWorksheet getExcelWorksheet (string filename)
        {
            FileInfo file = new FileInfo (filename);
            ExcelPackage excelPackage = new ExcelPackage (file);
            return excelPackage.Workbook.Worksheets.FirstOrDefault ();
        }

        string getMetadata ()
        {
                return excelWorksheet.Cells[4, 1].Value.ToString ();
        }

        DateTime parseDateTime (string rawDt)
        {
            string dtFormat = "dd-MM-yyyy";
            return DateTime.ParseExact (rawDt, dtFormat, CultureInfo.InvariantCulture);
        }
        DateTime getStartDate (string input)
        {
            return parseDateTime (input.Substring (input.IndexOf ("Data od:") + 9, 10));
        }

        DateTime getEndDate (string input)
        {
            return parseDateTime (input.Substring (input.IndexOf ("Data do:") + 9, 10));
        }

        string getUnit (string input)
        {
            return input.Substring (input.IndexOf ("Jednostka wykonująca:") + 22, 13).Trim (';');
        }

        List<MadeService> getMadeServices (string filename)
        {
            List<MadeService> madeServList = new List<MadeService> ();
            int rowsToCut = 4;
            int rowCount = excelWorksheet.Dimension.Rows - rowsToCut;
            int startRow = 6;

            for (int row = startRow; row <= rowCount; row++)
            {
                MadeService ms = new MadeService ();
                ms.Id = Int32.Parse (excelWorksheet.Cells[row, 1].Value.ToString ());
                ms.Date = parseDateTime (excelWorksheet.Cells[row, 2].Value.ToString ());
                ms.PatientName = excelWorksheet.Cells[row, 6].Value.ToString ();
                ms.PatientPesel = excelWorksheet.Cells[row, 8].Value.ToString ();
                ms.ServiceCode = ms.getServiceCode (excelWorksheet.Cells[row, 11].Value.ToString ());
                ms.Unit = ms.getUnit (excelWorksheet.Cells[row, 10].Value.ToString ());

                madeServList.Add (ms);
            }

            return madeServList;
        }
    }
}
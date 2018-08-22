using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;

namespace computerman_rtg_reports
{
    public class Report
    {
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public string Unit { get; set; }
        List<UnitReportItem> UnitReportItems { get; set; }
        List<ServiceReportItem> ServiceReportItems { get; set; }
        List<UnitServiceReportItem> UnitServiceReportItems { get; set; }
        RawUserData rawUserData;
        List<Service> pricelist;

        public Report (RawUserData rawUserData)
        {
            this.rawUserData = rawUserData;
            StartDate = getStartDate (rawUserData.Metadata);
            EndDate = getEndDate (rawUserData.Metadata);
            Unit = getUnit (rawUserData.Metadata);
            pricelist = getPricelist (Unit);
            UnitReportItems = getUnitsReport ();
            ServiceReportItems = getServiceReport ();
            UnitServiceReportItems = getUnitServiceReport ();
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
            input = input.Substring (input.IndexOf ("Jednostka wykonująca:") + 22, input.Length - input.IndexOf ("Jednostka wykonująca:") - 22);
            input = input.Substring (0, input.IndexOf (";"));
            return input;
        }

        List<Service> getPricelist (string unit)
        {
            return Service.getPricelist (unit);
        }

        List<UnitReportItem> getUnitsReport ()
        {
            List<UnitReportItem> unitReportItems = new List<UnitReportItem> ();

            try
            {
                foreach (var ms in rawUserData.MadeServicesList.GroupBy (p => p.Unit))
                {
                    UnitReportItem item = new UnitReportItem ();
                    item.Unit = ms.Key;
                    item.Count = ms.Count ();
                    item.Value = 0;

                    foreach (var serv in ms.Select (p => p.ServiceCode))
                    {
                        item.Value += pricelist.Single (p => p.Code == serv).Price;
                        if (Unit == "Pracownia RTG")
                        {
                            item.Photos += pricelist.Single (p => p.Code == serv).Photos;
                        }
                    }
                    unitReportItems.Add (item);
                }
            }
            catch (InvalidOperationException)
            {
                foreach (var ms in rawUserData.MadeServicesList)
                {
                    if (!pricelist.Exists (p => p.Code == ms.ServiceCode))
                    {
                        Console.WriteLine (ms.PatientName);
                    }
                }
            }

            return unitReportItems.OrderByDescending (p => p.Count).ToList ();
        }

        List<ServiceReportItem> getServiceReport ()
        {
            List<ServiceReportItem> serviceReportItems = new List<ServiceReportItem> ();

            foreach (var ms in rawUserData.MadeServicesList.GroupBy (p => p.ServiceCode))
            {
                ServiceReportItem item = new ServiceReportItem ();
                item.Code = ms.Key;
                item.Name = pricelist.Single (p => p.Code == item.Code).Name;
                item.Count = ms.Count ();
                item.Value = 0;

                foreach (var serv in ms.Select (p => p.ServiceCode))
                {
                    item.Value += pricelist.Single (p => p.Code == serv).Price;
                    if (Unit == "Pracownia RTG")
                    {
                        item.Photos += pricelist.Single (p => p.Code == serv).Photos;
                    }
                }
                serviceReportItems.Add (item);
            }

            return serviceReportItems.OrderByDescending (p => p.Count).ToList ();
        }

        List<UnitServiceReportItem> getUnitServiceReport ()
        {
            List<UnitServiceReportItem> services = new List<UnitServiceReportItem> ();

            int i = 1;
            foreach (var ms in rawUserData.MadeServicesList.GroupBy (p => p.Unit))
            {
                foreach (var serv in ms.GroupBy (p => p.ServiceCode))
                {
                    UnitServiceReportItem item = new UnitServiceReportItem ();
                    item.Unit = ms.Key;
                    item.Value = 0;
                    item.Id = i;
                    item.Code = serv.Key;
                    item.Count = ms.Where (p => p.ServiceCode == item.Code).Count ();
                    item.Value = pricelist.Single (p => p.Code == item.Code).Price * item.Count;
                    if (Unit == "Pracownia RTG")
                    {
                        item.Photos += pricelist.Single (p => p.Code == item.Code).Photos * item.Count;
                    }
                    services.Add (item);
                    i++;
                }
            }

            return services.OrderBy (p => p.Unit).ThenByDescending (p => p.Count).ToList ();
        }

        public string generate ()
        {
            string path = "Templates/report_rtg.xml";
            XDocument doc = XDocument.Load (path);
            XElement root = doc.Element ("Template");
            string output = string.Empty;

            if (Unit == "Pracownia RTG")
            {
                output = getRTGReport (root);
            }
            else
            {
                output = getReport (root);
            }

            output = output.Replace ("~^~^", "{{");
            output = output.Replace ("^~^~", "}}");
            output = output.Replace ("~^", "{");
            output = output.Replace ("^~", "}");

            string time = DateTime.Now.ToFileTime ().ToString ();

            string outputFile = getHash (output + time);
            File.WriteAllText ("tmp/" + outputFile + ".tex", output);

            Process process = new Process ();
            process.StartInfo.WorkingDirectory = "tmp";
            process.StartInfo.FileName = "pdflatex";
            process.StartInfo.Arguments = "-synctex=1 -interaction=nonstopmode " + outputFile + ".tex";
            process.Start ();
            process.Dispose ();
            return outputFile + ".pdf";
        }

        string getReport (XElement root)
        {
            string header = root.Element ("Header").Value;

            // units report
            string report1Header = root.Element ("Report1Header").Value;
            report1Header = string.Format (report1Header, StartDate.ToShortDateString (), EndDate.ToShortDateString (), Unit);

            string table1Header = root.Element ("Table1Header").Value;
            string table1Item = root.Element ("Table1Row").Value;
            int i = 1;
            int report1Count = 0;
            decimal report1Value = 0;
            foreach (UnitReportItem item in UnitReportItems)
            {
                string newItem = string.Format (table1Item, i, item.Unit, item.Count, item.Value.ToString ("c"));
                table1Header += newItem;
                report1Count += item.Count;
                report1Value += item.Value;
                i++;
            }

            string table1Summary = root.Element ("Table1Summary").Value;
            table1Summary = string.Format (table1Summary, report1Count, report1Value.ToString ("c"));

            string report1 = report1Header + table1Header + table1Summary;

            // services report
            string report2Header = root.Element ("Report2Header").Value;
            report2Header = string.Format (report2Header, StartDate.ToShortDateString (), EndDate.ToShortDateString (), Unit);

            string table2Header = root.Element ("Table2Header").Value;
            string table2Item = root.Element ("Table2Row").Value;
            i = 1;
            int report2Count = 0;
            decimal report2Value = 0;
            foreach (ServiceReportItem item in ServiceReportItems)
            {
                string newItem = string.Format (table2Item, i, item.Code, item.Name, item.Count, item.Value.ToString ("c"));
                table2Header += newItem;
                report2Count += item.Count;
                report2Value += item.Value;
                i++;
            }

            string table2Summary = root.Element ("Table2Summary").Value;
            table2Summary = string.Format (table2Summary, report2Count, report2Value.ToString ("c"));

            string report2 = report2Header + table2Header + table2Summary;

            // unitservice report
            string report3Header = root.Element ("Report3Header").Value;
            report3Header = string.Format (report3Header, StartDate.ToShortDateString (), EndDate.ToShortDateString (), Unit);

            string table3Header = root.Element ("Table3Header").Value;
            string table3Item = root.Element ("Table3Row").Value;
            i = 1;
            int report3Count = 0;
            decimal report3Value = 0;
            foreach (UnitServiceReportItem item in UnitServiceReportItems)
            {
                string newItem = string.Format (table3Item, i, item.Unit, item.Code, item.Count, item.Value.ToString ("c"));
                table3Header += newItem;
                report3Count += item.Count;
                report3Value += item.Value;
                i++;
            }

            string table3Summary = root.Element ("Table3Summary").Value;
            table3Summary = string.Format (table3Summary, report3Count, report3Value.ToString ("c"));

            string report3 = report3Header + table3Header + table3Summary;

            string output = header + report1 + report2 + report3;
            return output;
        }

        string getRTGReport (XElement root)
        {
            string header = root.Element ("Header").Value;

            // units report
            string report1Header = root.Element ("Report1Header").Value;
            report1Header = string.Format (report1Header, StartDate.ToShortDateString (), EndDate.ToShortDateString (), Unit);

            string table1Header = root.Element ("Table1HeaderRTG").Value;
            string table1Item = root.Element ("Table1RowRTG").Value;
            int i = 1;
            int report1Count = 0;
            int report1Photos = 0;
            decimal report1Value = 0;
            foreach (UnitReportItem item in UnitReportItems)
            {
                string newItem = string.Format (table1Item, i, item.Unit, item.Count, item.Photos, item.Value.ToString ("c"));
                table1Header += newItem;
                report1Count += item.Count;
                report1Photos += item.Photos;
                report1Value += item.Value;
                i++;
            }

            string table1Summary = root.Element ("Table1SummaryRTG").Value;
            table1Summary = string.Format (table1Summary, report1Count, report1Photos, report1Value.ToString ("c"));

            string report1 = report1Header + table1Header + table1Summary;

            // services report
            string report2Header = root.Element ("Report2Header").Value;
            report2Header = string.Format (report2Header, StartDate.ToShortDateString (), EndDate.ToShortDateString (), Unit);

            string table2Header = root.Element ("Table2HeaderRTG").Value;
            string table2Item = root.Element ("Table2RowRTG").Value;
            i = 1;
            int report2Count = 0;
            int report2Photos = 0;
            decimal report2Value = 0;
            foreach (ServiceReportItem item in ServiceReportItems)
            {
                string newItem = string.Format (table2Item, i, item.Code, item.Name, item.Count, item.Photos, item.Value.ToString ("c"));
                table2Header += newItem;
                report2Count += item.Count;
                report2Photos += item.Photos;
                report2Value += item.Value;
                i++;
            }

            string table2Summary = root.Element ("Table2SummaryRTG").Value;
            table2Summary = string.Format (table2Summary, report2Count, report2Photos, report2Value.ToString ("c"));

            string report2 = report2Header + table2Header + table2Summary;

            // unitservice report
            string report3Header = root.Element ("Report3Header").Value;
            report3Header = string.Format (report3Header, StartDate.ToShortDateString (), EndDate.ToShortDateString (), Unit);

            string table3Header = root.Element ("Table3HeaderRTG").Value;
            string table3Item = root.Element ("Table3RowRTG").Value;
            i = 1;
            int report3Count = 0;
            int report3Photos = 0;
            decimal report3Value = 0;
            foreach (UnitServiceReportItem item in UnitServiceReportItems)
            {
                string newItem = string.Format (table3Item, i, item.Unit, item.Code, item.Count, item.Photos, item.Value.ToString ("c"));
                table3Header += newItem;
                report3Count += item.Count;
                report3Photos += item.Photos;
                report3Value += item.Value;
                i++;
            }

            string table3Summary = root.Element ("Table3SummaryRTG").Value;
            table3Summary = string.Format (table3Summary, report3Count, report3Photos, report3Value.ToString ("c"));

            string report3 = report3Header + table3Header + table3Summary;

            string output = header + report1 + report2 + report3;
            return output;
        }

        string getHash (string input)
        {
            string hashAlgo = "SHA256";
            HashAlgorithm algo = HashAlgorithm.Create (hashAlgo);
            byte[] hashBytes = algo.ComputeHash (Encoding.UTF8.GetBytes (input));

            StringBuilder sb = new StringBuilder ();
            foreach (byte b in hashBytes)
            {
                sb.Append (b.ToString ("X2"));
            }
            string computedHash = sb.ToString ();

            return computedHash;
        }
    }
}
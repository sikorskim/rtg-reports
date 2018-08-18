using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using computerman_rtg_reports.Models;
using Microsoft.AspNetCore.Http;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using OfficeOpenXml;
using System.Globalization;

namespace computerman_rtg_reports.Controllers
{
    public class ReportController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Index(IFormFile file)
        {
            var fileName = @"tmp/" + getHash(DateTime.Now.ToLongTimeString());// + ".xlsx";

            using (var stream = new FileStream(fileName, FileMode.Create))
            {
                //file.CopyTo(stream);
                await file.CopyToAsync(stream);
            }

            // process uploaded files
            // Don't rely on or trust the FileName property without validation.

            //return RedirectToAction("Import", new { filename = fileName });
            return RedirectToAction("GenerateReport3", new { filename = fileName });
        }

        public IActionResult GenerateReport(string filename)
        {
            FileInfo file = new FileInfo(filename);
            List<MadeService> madeServList = new List<MadeService>();

            // try
            // {
            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();
                int rowCount = worksheet.Dimension.Rows;
                int startRow = 2;
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

            UnitsReport unitsReport = new UnitsReport();
            unitsReport.StartDate = DateTime.Now;
            unitsReport.EndDate = DateTime.Now;
            unitsReport.Unit = "Pracownia USG";
            unitsReport.Items = new List<UnitsReportItem>();

            List<Service> pricelist = Service.getPricelist("usg");

            foreach (var ms in madeServList.GroupBy(p => p.Unit))
            {
                UnitsReportItem item = new UnitsReportItem();
                item.Unit = ms.Key;
                //item.Count=madeServList.Where(p=>p.Unit==item.Unit).Count();
                item.Count = ms.Count();
                item.Value = 0;

                foreach (var serv in ms.Select(p => p.ServiceCode))
                {
                    item.Value += pricelist.Single(p => p.Code == serv).Price;
                }
                unitsReport.Items.Add(item);
            }

            List<UnitsReportItem> orderedList = unitsReport.Items.OrderByDescending(p => p.Count).ToList();
            return View(orderedList);
        }

        public IActionResult GenerateReport2(string filename)
        {
            FileInfo file = new FileInfo(filename);
            List<MadeService> madeServList = new List<MadeService>();

            // try
            // {
            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();
                int rowCount = worksheet.Dimension.Rows;
                int startRow = 2;
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

            List<ServiceReportItem> services = new List<ServiceReportItem>();

            List<Service> pricelist = Service.getPricelist("usg");

            foreach (var ms in madeServList.GroupBy(p => p.ServiceCode))
            {
                ServiceReportItem item = new ServiceReportItem();
                item.Code = ms.Key;
                item.Name = pricelist.Single(p=>p.Code==item.Code).Name;
                item.Count = ms.Count();
                item.Value = 0;

                foreach (var serv in ms.Select(p => p.ServiceCode))
                {
                    item.Value += pricelist.Single(p => p.Code == serv).Price;
                }
                services.Add(item);
            }

            List<ServiceReportItem> orderedList = services.OrderByDescending(p => p.Count).ToList();
            return View(orderedList);
        }

        public IActionResult GenerateReport3(string filename)
        {
            FileInfo file = new FileInfo(filename);
            List<MadeService> madeServList = new List<MadeService>();

            // try
            // {
            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();
                int rowCount = worksheet.Dimension.Rows;
                int startRow = 2;
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

            List<UnitServiceReportItem> services = new List<UnitServiceReportItem>();

            List<Service> pricelist = Service.getPricelist("usg");
            int i =1;
            foreach (var ms in madeServList.GroupBy(p => p.Unit))
            {
                UnitServiceReportItem item = new UnitServiceReportItem();                
                item.Unit=ms.Key;
                item.Value = 0;

                foreach (var serv in ms.GroupBy(p=>p.ServiceCode))
                {
                    item.Id=i;
                    item.Code = serv.Key;
                    item.Count = ms.Where(p=>p.ServiceCode==item.Code).Count();
                    item.Value = pricelist.Single(p => p.Code == item.Code).Price*item.Count;
                    services.Add(item);
                    i++;
                }                
            }

            List<UnitServiceReportItem> orderedList = services.OrderByDescending(p => p.Count).ToList();
            return View(orderedList);
        }

        public IActionResult Import(string filename)
        {
            FileInfo file = new FileInfo(filename);
            List<MadeService> madeServList = new List<MadeService>();

            // try
            // {
            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();
                int rowCount = worksheet.Dimension.Rows;
                int startRow = 2;
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
                //return View(madeServList);
                return RedirectToAction("GenerateReport", new { madeServLst = madeServList });
            }
        }

        string getHash(string input)
        {
            string hashAlgo = "SHA256";
            HashAlgorithm algo = HashAlgorithm.Create(hashAlgo);
            byte[] hashBytes = algo.ComputeHash(Encoding.UTF8.GetBytes(input));

            StringBuilder sb = new StringBuilder();
            foreach (byte b in hashBytes)
            {
                sb.Append(b.ToString("X2"));
            }
            string computedHash = sb.ToString();

            return computedHash;
        }
    }
}

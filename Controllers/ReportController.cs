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

            return RedirectToAction("Import", new { filename = fileName });
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
                    ms.Unit = ms.getUnit(worksheet.Cells[row,10].Value.ToString());

                    madeServList.Add(ms);
                }
                return View(madeServList);
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

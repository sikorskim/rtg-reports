using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using computerman_rtg_reports.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;

namespace computerman_rtg_reports.Controllers
{
    public class ReportController : Controller
    {
        public IActionResult Index ()
        {
            return View ();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Index (IFormFile file)
        {
            var fileName = "tmp/" + getHash (DateTime.Now.ToLongTimeString ());

            using (var stream = new FileStream (fileName, FileMode.Create))
            {
                //file.CopyTo(stream);
                await file.CopyToAsync (stream);
            }
            
            RawUserData rawUserData = new RawUserData(fileName);
            Report report = new Report(rawUserData);            
            string pdfFilename = report.generate();
            await Task.Delay (1000);           

            return RedirectToAction ("GetPdfFile", new { filename = pdfFilename});
        }

        public IActionResult GetPdfFile (string filename)
        {
            const string contentType = "application/pdf";
            HttpContext.Response.ContentType = contentType;
            FileContentResult result = null;

            try
            {
                result = new FileContentResult (System.IO.File.ReadAllBytes (filename), contentType)
                {
                    FileDownloadName = "report.pdf"
                };
                //deleteTempFiles (filename.Substring (4, 64));
                return result;
            }
            catch (FileNotFoundException)
            {
                return NotFound ();
            }
        }

        void deleteTempFiles(string filename)
        {
            DirectoryInfo dir = new DirectoryInfo("tmp");
            foreach (FileInfo file in dir.GetFiles())
            {
                if (file.Name.Contains(filename))
                {
                    file.Delete();
                }
            }
        }

        public IActionResult Import (string filename)
        {
            FileInfo file = new FileInfo (filename);
            List<MadeService> madeServList = new List<MadeService> ();

            using (ExcelPackage package = new ExcelPackage (file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault ();
                int rowCount = worksheet.Dimension.Rows;
                int startRow = 2;
                string dtFormat = "dd-MM-yyyy";

                for (int row = startRow; row <= rowCount; row++)
                {
                    MadeService ms = new MadeService ();
                    ms.Id = Int32.Parse (worksheet.Cells[row, 1].Value.ToString ());
                    ms.Date = DateTime.ParseExact (worksheet.Cells[row, 2].Value.ToString (), dtFormat, CultureInfo.InvariantCulture);
                    ms.PatientName = worksheet.Cells[row, 6].Value.ToString ();
                    ms.PatientPesel = worksheet.Cells[row, 8].Value.ToString ();
                    ms.ServiceCode = ms.getServiceCode (worksheet.Cells[row, 11].Value.ToString ());
                    ms.Unit = ms.getUnit (worksheet.Cells[row, 10].Value.ToString ());

                    madeServList.Add (ms);
                }
                //return View(madeServList);
                return RedirectToAction ("GenerateReport", new { madeServLst = madeServList });
            }
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
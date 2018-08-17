using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using computerman_rtg_reports.Models;
using System.Xml.Linq;
using Microsoft.AspNetCore.Hosting;
using System.IO;
using OfficeOpenXml;
using System.Text;
using System.Globalization;

namespace computerman_rtg_reports.Controllers
{
    public class ServiceController : Controller
    {
        public IActionResult Index()
        {
            string path = "Pricelists/usg.xml";
            XDocument doc = XDocument.Load(path);
            XElement root = doc.Element("Pricelist");

            ViewData["PricelistName"]=root.Attribute("Name").Value;
            List<Service> services = new List<Service>();

            foreach(XElement elem in root.Elements("Item"))
            {
                Service serv = new Service();
                serv.Id=Int32.Parse(elem.Attribute("Id").Value);
                serv.Code=elem.Attribute("Code").Value;
                serv.Name=elem.Attribute("Name").Value;
                serv.Price=decimal.Parse(elem.Attribute("Price").Value);

                services.Add(serv);
            }

            return View(services);
        }

public IActionResult Import()
{
    string sFileName = @"tmp/usg0618.xlsx";
    FileInfo file = new FileInfo(sFileName);
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
                ms.Id=Int32.Parse(worksheet.Cells[row, 1].Value.ToString());
                ms.Date=DateTime.ParseExact(worksheet.Cells[row,2].Value.ToString(), dtFormat, CultureInfo.InvariantCulture);
                ms.PatientName=worksheet.Cells[row,6].Value.ToString();
                ms.PatientPesel=worksheet.Cells[row,8].Value.ToString();
                ms.ServiceCode = ms.getServiceCode(worksheet.Cells[row,11].Value.ToString());

                madeServList.Add(ms);
            }
            return View(madeServList);
        }
    // }
    // catch (Exception ex)
    // {
    //     return View("Some error occured while importing." + ex.Message);
    // }
}
    }
}
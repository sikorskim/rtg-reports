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

namespace computerman_rtg_reports.Controllers
{
    public class ServiceController : Controller
    {
        public IActionResult Index()
        {
            string path = "pricelists/usg.xml";
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
    // try
    // {
        using (ExcelPackage package = new ExcelPackage(file))
        {
            StringBuilder sb = new StringBuilder();
            ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();
            int rowCount = worksheet.Dimension.Rows;
            int ColCount = worksheet.Dimension.Columns;
            bool bHeaderRow = true;
            for (int row = 1; row <= rowCount; row++)
            {
                for (int col = 1; col <= ColCount; col++)
                {
                    if (bHeaderRow)
                    {
                        sb.Append(worksheet.Cells[row, col].Value.ToString() + "\t");
                    }
                    else
                    {
                        sb.Append(worksheet.Cells[row, col].Value.ToString() + "\t");
                    }
                }
                sb.Append(Environment.NewLine);
            }
            return View(sb.ToString());
        }
    // }
    // catch (Exception ex)
    // {
    //     return View("Some error occured while importing." + ex.Message);
    // }
}
    }
}
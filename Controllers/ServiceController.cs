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

            ViewData["PricelistName"] = root.Attribute("Name").Value;
            List<Service> services = new List<Service>();

            foreach (XElement elem in root.Elements("Item"))
            {
                Service serv = new Service();
                serv.Id = Int32.Parse(elem.Attribute("Id").Value);
                serv.Code = elem.Attribute("Code").Value;
                serv.Name = elem.Attribute("Name").Value;
                serv.Price = decimal.Parse(elem.Attribute("Price").Value);

                services.Add(serv);
            }

            return View(services);
        }
    }
}
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using computerman_rtg_reports.Models;
using Microsoft.AspNetCore.Mvc.Rendering;

namespace computerman_rtg_reports.Controllers
{
    public class ServiceController : Controller
    {
        public IActionResult Index()
        {
            Dictionary<int, string> pricelists = new Dictionary<int, string>()
            {
                {0, "USG"},
                {1, "RTG"},
                {2, "KT"}
            };
            ViewData["Pricelist"] = new SelectList (pricelists, "Key", "Value", 0);
            return View(Service.getPricelist("usg"));
        }
    }
}
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using computerman_rtg_reports.Models;

namespace computerman_rtg_reports.Controllers
{
    public class ServiceController : Controller
    {
        public IActionResult Index()
        {
            return View(Service.getPricelist("usg"));
        }
    }
}
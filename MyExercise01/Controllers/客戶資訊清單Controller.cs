using MyExercise01.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace MyExercise01.Controllers
{
    public class 客戶資訊清單Controller : Controller
    {
        private CustomerInfoEntities db = new CustomerInfoEntities();
        // GET: 客戶資訊清單
        public ActionResult Index()
        {
            return View(db.View_客戶資訊清單.ToList());
        }
    }
}
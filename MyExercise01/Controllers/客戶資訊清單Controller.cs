using MyExercise01.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace MyExercise01.Controllers
{
    public class 客戶資訊清單Controller : BaseController
    {
        private CustomerInfoEntities db = new CustomerInfoEntities();
        // GET: 客戶資訊清單
        public ActionResult Index()
        {
            return View(repoView_客戶資訊清單.All().ToList());
        }

        [HttpPost]
        public ActionResult Index(string button)
        {
            var data = repoView_客戶資訊清單.All();
            return File(repoView_客戶資訊清單.ExprotExcel(data.ToList()),
                        "application / vnd.openxmlformats - officedocument.spreadsheetml.sheet", "Export.xlsx");
        }
    }
}
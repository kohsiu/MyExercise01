using MyExercise01.Models;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web.Mvc;
using System;
using System.IO;
using NPOI.XSSF.UserModel;

namespace MyExercise01.Controllers
{
    public class 客戶資料Controller : BaseController
    {
        // GET: 客戶資料
        public ActionResult Index()
        {
            var data = repo客戶資料.All();
            ViewBag.客戶分類 = new SelectList(repo客戶資料.CreateCusKindData(), "Kind", "Kind");
            return View(data.ToList());
        }

        [HttpPost]
        public ActionResult Index(string keyword, string 客戶分類, string button)
        {

            var data = repo客戶資料.SearchCusName(keyword, 客戶分類);
            ViewBag.客戶分類 = new SelectList(repo客戶資料.CreateCusKindData(), "Kind", "Kind");

            switch (button)
            {
                case "Search":
                    return View(data.ToList());

                case "Excel":
                    return File(repo客戶資料.ExprotExcel(data.ToList()),
                        "application / vnd.openxmlformats - officedocument.spreadsheetml.sheet", "Export.xlsx");
            }
            return View(data.ToList());
        }

        // GET: 客戶資料/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }

            客戶資料 客戶資料 = repo客戶資料.Find(id.Value);
            if (客戶資料 == null)
            {
                return HttpNotFound();
            }
            return View(客戶資料);
        }

        [HttpPost]
        public ActionResult Details(IList<BatchUpdate客戶聯絡人> data)
        {
            if (ModelState.IsValid)
            {
                foreach (var item in data)
                {
                    var 客戶聯絡人 = repo客戶聯絡人.Find(item.Id);

                    客戶聯絡人.職稱 = item.職稱;
                    客戶聯絡人.手機 = item.手機;
                    客戶聯絡人.電話 = item.電話;
                }

                repo客戶聯絡人.UnitOfWork.Commit();

                return RedirectToAction("Index");
            }

            ViewData.Model = repo客戶資料.Find(data.FirstOrDefault().客戶Id);
            return View();
        }

        // GET: 客戶資料/Create
        public ActionResult Create()
        {
            ViewBag.客戶分類 = new SelectList(repo客戶資料.CreateCusKindData(), "Kind", "Kind");
            return View();
        }

        // POST: 客戶資料/Create
        // 若要免於過量張貼攻擊，請啟用想要繫結的特定屬性，如需
        // 詳細資訊，請參閱 http://go.microsoft.com/fwlink/?LinkId=317598。
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "Id,客戶名稱,統一編號,電話,傳真,地址,Email,客戶分類,是否已刪除")] 客戶資料 客戶資料)
        {
            if (ModelState.IsValid)
            {
                repo客戶資料.Add(客戶資料);
                repo客戶資料.UnitOfWork.Commit();
                return RedirectToAction("Index");
            }

            ViewBag.客戶分類 = new SelectList(repo客戶資料.CreateCusKindData(), "Kind", "Kind", 客戶資料.客戶分類);
            return View(客戶資料);
        }



        // GET: 客戶資料/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }

            客戶資料 客戶資料 = repo客戶資料.Find(id.Value);
            if (客戶資料 == null)
            {
                return HttpNotFound();
            }
            ViewBag.客戶分類 = new SelectList(repo客戶資料.CreateCusKindData(), "Kind", "Kind", 客戶資料.客戶分類);
            return View(客戶資料);
        }

        // POST: 客戶資料/Edit/5
        // 若要免於過量張貼攻擊，請啟用想要繫結的特定屬性，如需
        // 詳細資訊，請參閱 http://go.microsoft.com/fwlink/?LinkId=317598。
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "Id,客戶名稱,統一編號,電話,傳真,地址,Email,客戶分類,是否已刪除")] 客戶資料 客戶資料)
        {
            if (ModelState.IsValid)
            {
                repo客戶資料.UnitOfWork.Context.Entry(客戶資料).State = EntityState.Modified;
                repo客戶資料.UnitOfWork.Commit();
                return RedirectToAction("Index");
            }
            ViewBag.客戶分類 = new SelectList(repo客戶資料.CreateCusKindData(), "Kind", "Kind", 客戶資料.客戶分類);
            return View(客戶資料);
        }

        // GET: 客戶資料/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }

            客戶資料 客戶資料 = repo客戶資料.Find(id.Value);
            if (客戶資料 == null)
            {
                return HttpNotFound();
            }
            return View(客戶資料);
        }

        // POST: 客戶資料/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            客戶資料 客戶資料 = repo客戶資料.Find(id);
            repo客戶資料.Delete(客戶資料);
            //foreach (var cusContact in 客戶資料.客戶聯絡人)
            //{
            //    cusContact.是否已刪除 = true;
            //}

            //foreach (var bank in 客戶資料.客戶銀行資訊)
            //{
            //    bank.是否已刪除 = true;
            //}

            //db.客戶資料.Remove(客戶資料);

            repo客戶資料.UnitOfWork.Commit();
            return RedirectToAction("Index");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                repo客戶資料.UnitOfWork.Context.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using AnkietaProjekt.DAL;
using AnkietaProjekt.Models;

namespace AnkietaProjekt.Controllers
{
    public class TestsController : Controller
    {
        private QuestionContext db = new QuestionContext();

        // GET: Tests
        public ActionResult Index()
        {
            var tests = from a in db.Testy select a;
            if (User.IsInRole("Admin"))
            {
                return View(tests.ToList());
            }
            else
            {
                return View(tests.Where(o => o.Username == User.Identity.Name).ToList());
            }
            
        }

        // GET: Tests/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Test test = db.Testy.Find(id);
            if (test == null)
            {
                return HttpNotFound();
            }
            return View(test);
        }

        // GET: Tests/Create
        public ActionResult Create()
        {
            Test test = new Test();
            List<String> quest = new List<String>();
            foreach (Question q in db.Questions.ToList())
            {
                quest.Add(q.Pytanie);
            }
            test.pytanie1 = quest[0];
            test.pytanie2 = quest[1];
            test.pytanie3 = quest[2];
            test.pytanie4 = quest[3];
            test.pytanie5 = quest[4];
            test.pytanie6 = quest[5];
            test.pytanie7 = quest[6];
            test.pytanie8 = quest[7];
            test.pytanie9 = quest[8];
            test.pytanie10 = quest[9];
            test.pytanie11 = quest[10];
            test.pytanie12 = quest[11];
            test.pytanie13 = quest[12];
            test.pytanie14 = quest[13];
            test.pytanie15 = quest[14];
            test.pytanie16 = quest[15];
            test.pytanie17 = quest[16];
            test.pytanie18 = quest[17];
            test.pytanie19 = quest[18];
            test.pytanie20 = quest[19];
            test.Username = User.Identity.Name;

            return View(test);
        }

        // POST: Tests/Create
        // Aby zapewnić ochronę przed atakami polegającymi na przesyłaniu dodatkowych danych, włącz określone właściwości, z którymi chcesz utworzyć powiązania.
        // Aby uzyskać więcej szczegółów, zobacz https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "ID,Username,plec,pytanie1,odp1,pytanie2,odp2,pytanie3,odp3,pytanie4,odp4,pytanie5,odp5,pytanie6,odp6,pytanie7,odp7,pytanie8,odp8,pytanie9,odp9,pytanie10,odp10,pytanie11,odp11,pytanie12,odp12,pytanie13,odp13,pytanie14,odp14,pytanie15,odp15,pytanie16,odp16,pytanie17,odp17,pytanie18,odp18,pytanie19,odp119,pytanie20,odp20")] Test test)
        {
            if (ModelState.IsValid)
            {
            
                List<String> quest = new List<String>();
                foreach (Question q in db.Questions.ToList())
                {
                    quest.Add(q.Pytanie);
                }
                test.pytanie1 = quest[0];
                test.pytanie2 = quest[1];
                test.pytanie3 = quest[2];
                test.pytanie4 = quest[3];
                test.pytanie5 = quest[4];
                test.pytanie6 = quest[5];
                test.pytanie7 = quest[6];
                test.pytanie8 = quest[7];
                test.pytanie9 = quest[8];
                test.pytanie10 = quest[9];
                test.pytanie11 = quest[10];
                test.pytanie12 = quest[11];
                test.pytanie13 = quest[12];
                test.pytanie14 = quest[13];
                test.pytanie15 = quest[14];
                test.pytanie16 = quest[15];
                test.pytanie17 = quest[16];
                test.pytanie18 = quest[17];
                test.pytanie19 = quest[18];
                test.pytanie20 = quest[19];
                test.Username = User.Identity.Name;
                db.Testy.Add(test);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(test);
        }

        // GET: Tests/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Test test = db.Testy.Find(id);
            if (test == null)
            {
                return HttpNotFound();
            }
            return View(test);
        }

        // POST: Tests/Edit/5
        // Aby zapewnić ochronę przed atakami polegającymi na przesyłaniu dodatkowych danych, włącz określone właściwości, z którymi chcesz utworzyć powiązania.
        // Aby uzyskać więcej szczegółów, zobacz https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "ID,Username,plec,pytanie1,odp1,pytanie2,odp2,pytanie3,odp3,pytanie4,odp4,pytanie5,odp5,pytanie6,odp6,pytanie7,odp7,pytanie8,odp8,pytanie9,odp9,pytanie10,odp10,pytanie11,odp11,pytanie12,odp12,pytanie13,odp13,pytanie14,odp14,pytanie15,odp15,pytanie16,odp16,pytanie17,odp17,pytanie18,odp18,pytanie19,odp119,pytanie20,odp20")] Test test)
        {
            if (ModelState.IsValid)
            {
                db.Entry(test).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(test);
        }

        // GET: Tests/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Test test = db.Testy.Find(id);
            if (test == null)
            {
                return HttpNotFound();
            }
            return View(test);
        }

        // POST: Tests/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            Test test = db.Testy.Find(id);
            db.Testy.Remove(test);
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}

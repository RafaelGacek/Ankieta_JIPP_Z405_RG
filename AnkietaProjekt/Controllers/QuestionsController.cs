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
using _Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace AnkietaProjekt.Controllers
{
    public class QuestionsController : Controller
    {
        private QuestionContext db = new QuestionContext();

        // GET: Questions
        public ActionResult Index()
        {
            return View(db.Questions.ToList());
        }

        // GET: Questions/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Question question = db.Questions.Find(id);
            if (question == null)
            {
                return HttpNotFound();
            }
            return View(question);
        }
        [HttpGet]
        public ActionResult OpenFile()
        {
            return View();
        }

        [HttpPost]
        public ActionResult OpenFile(HttpPostedFileBase excelFile)
        {
            try
            {
                string filename = excelFile.FileName;
                //Path.GetExtension(excelFile.FileName);
                if (excelFile != null && (filename.EndsWith(".xls")))
                {
                    string path = Server.MapPath("~/Content/") + Guid.NewGuid() + filename;
                    excelFile.SaveAs(path);
                    Spire.Xls.Workbook workbook = new Spire.Xls.Workbook();
                    workbook.LoadFromFile($@"{path}");
                    Spire.Xls.Worksheet sheet = workbook.Worksheets[0];
                    string FinalPath = HttpContext.Server.MapPath("~/Files/");
                    sheet.SaveToFile($@"{FinalPath}" + $@"{filename.Replace(".xls", ".csv")}", ",", System.Text.Encoding.UTF8);

                    StreamReader reader = null;
                    int lncnt = 0;
                    List<string> line;
                    string csv = filename.Replace(".xls", ".csv");
                    string pathDownload = FinalPath + csv;
                    reader = new StreamReader(Path.Combine(pathDownload), System.Text.Encoding.UTF8);
                    string header = reader.ReadLine();
                    string hd = @"Pytania";
                    List<Question> questions = new List<Question>();
                    if (header.Replace(" ", "") == hd.Replace(" ", ""))
                    {
                        while (!reader.EndOfStream)
                        {
                            lncnt++;
                            line = reader.ReadLine().Split(',').Select(t => t.Trim('"', '\'')).ToList();
                            Question question = new Question();
                            question.Pytanie = line[0];
                            questions.Add(question);

                        }
                        foreach (var item in questions)
                        {
                            db.Questions.Add(item);
                        }


                        db.SaveChanges();

                        reader.Close();

                        Files f = new Files(FinalPath + filename.Replace(".xls", ".csv"));
                        f.RemoveFile();
                        return RedirectToAction("Index");
                    }


                }
            }
            catch (FileNotFoundException e)
            {
                Console.WriteLine("Nie znaleziono pliku" + e);
            }
            // var parti = from a in db.ShopTable select a;
            // return View(parti.ToList());
            return RedirectToAction("Index");
        }


        // GET: Questions/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: Questions/Create
        // Aby zapewnić ochronę przed atakami polegającymi na przesyłaniu dodatkowych danych, włącz określone właściwości, z którymi chcesz utworzyć powiązania.
        // Aby uzyskać więcej szczegółów, zobacz https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "ID,Pytanie")] Question question)
        {
            if (ModelState.IsValid)
            {
                db.Questions.Add(question);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(question);
        }

        // GET: Questions/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Question question = db.Questions.Find(id);
            if (question == null)
            {
                return HttpNotFound();
            }
            return View(question);
        }

        // POST: Questions/Edit/5
        // Aby zapewnić ochronę przed atakami polegającymi na przesyłaniu dodatkowych danych, włącz określone właściwości, z którymi chcesz utworzyć powiązania.
        // Aby uzyskać więcej szczegółów, zobacz https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "ID,Pytanie")] Question question)
        {
            if (ModelState.IsValid)
            {
                db.Entry(question).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(question);
        }

        // GET: Questions/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Question question = db.Questions.Find(id);
            if (question == null)
            {
                return HttpNotFound();
            }
            return View(question);
        }

        // POST: Questions/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            Question question = db.Questions.Find(id);
            db.Questions.Remove(question);
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

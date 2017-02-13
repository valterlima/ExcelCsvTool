using CsvConverter.Controllers.src;
using System;
using System.IO;
using System.Web;
using System.Web.Mvc;

namespace CsvConverter.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }
        [HttpGet]
        public ActionResult UploadFile()
        {
            return RedirectToAction("Index");
        }

        [HttpPost]
        public ActionResult UploadFile(HttpPostedFileBase file)
        {
            try
            {
                CsvManager manager = new CsvManager(file);

                string downloadPath = manager.GenerateDownloadPath();

                //zip it all or not

                ViewBag.Message = "The process is complete! Here's your download link: ";
                ViewBag.DownloadLink = Url.Encode(downloadPath);
                return View("Index");
            }
            catch (Exception ex)
            {
                ViewBag.Message = "File upload failed! " + ex.Message;
                return View("Index");
            }
        }

        public ActionResult Download(string file)
        {
            var path = Path.Combine(Server.MapPath("~/UploadedFiles"),  file);
            if (!System.IO.File.Exists(path))
            {
                return HttpNotFound();
            }

            string downloadName = new FileInfo(path).Name;

            var fileBytes = System.IO.File.ReadAllBytes(path);
            var response = new FileContentResult(fileBytes, "application/octet-stream")
            {
                FileDownloadName = downloadName
            };
            return response;
        }
    }
}
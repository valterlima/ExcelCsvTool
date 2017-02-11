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

        [HttpPost]
        public ActionResult UploadFile(HttpPostedFileBase file)
        {
            try
            {
                CsvManager manager = new CsvManager(file);

                string zipFile = manager.zipFilePath;

                //zip it all or not

                ViewBag.Message = "The process is complete! Here's your download link: ";
                ViewBag.DownloadLink = zipFile;
                return View("Index");
            }
            catch (Exception ex)
            {
                ViewBag.Message = "File upload failed!! Exception: " + ex.Message;
                return View("Index");
            }
        }

        public ActionResult Download(string file)
        {
            if (!System.IO.File.Exists(file))
            {
                return HttpNotFound();
            }

            string downloadName = new FileInfo(file).Name;

            var fileBytes = System.IO.File.ReadAllBytes(file);
            var response = new FileContentResult(fileBytes, "application/octet-stream")
            {
                FileDownloadName = new FileInfo(file).Name
            };
            return response;
        }
    }
}
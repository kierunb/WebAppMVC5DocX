using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Xml.Linq;

namespace WebAppMVC5DocX.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult Docx()
        {
            string filePath = Server.MapPath(@"~/Documents/test-docx.docx");
            
            byte[] byteArray = System.IO.File.ReadAllBytes(filePath);

            using (var memoryStream = new MemoryStream())
            {
                memoryStream.Write(byteArray, 0, byteArray.Length);
                using (var doc = WordprocessingDocument.Open(memoryStream, true))
                {
                    var settings = new HtmlConverterSettings()
                    {
                        PageTitle = "My Page Title"
                    };
                    var html = HtmlConverter.ConvertToHtml(doc, settings);

                    ViewBag.HtmlOutput = html.ToString();

                    //File.WriteAllText(HTMLFilePath, html.ToStringNewLineOnAttributes());
                }
            }


            return View();
        }




        public ActionResult Pdf()
        {

            return View();
        }

        public ActionResult GetPdf()
        {

            string filePath = Server.MapPath(@"~/Documents/test-pdf.pdf");

            byte[] byteArray = System.IO.File.ReadAllBytes(filePath);

            return File(byteArray, "application/pdf");
        }
        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}
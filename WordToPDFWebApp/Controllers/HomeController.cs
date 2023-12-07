using Microsoft.AspNetCore.Mvc;
using Syncfusion.Pdf;
using Syncfusion.Pdf.Graphics;
using Syncfusion.Drawing;
using System.IO;
using System.Diagnostics;
using WordToPDFWebApp.Models;
using Microsoft.Extensions.Hosting.Internal;
using Microsoft.AspNetCore.Hosting.Server;
using System.Reflection;

namespace WordToPDFWebApp.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private Microsoft.AspNetCore.Hosting.IHostingEnvironment _env;
        public HomeController(Microsoft.AspNetCore.Hosting.IHostingEnvironment env)
        {
            _env = env;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
        public IActionResult WordToPDF(string button)
        {
            try
            {
                string fontstring = Environment.CurrentDirectory + "\\arial.ttf";
                //Create a new PDF document.
                PdfDocument document = new PdfDocument();
                //Add a page to the document.
                PdfPage page = document.Pages.Add();

                //Create PDF graphics for the page.
                PdfGraphics graphics = page.Graphics;
                FileStream fontStream = new FileStream(fontstring, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.ReadWrite);
                //Use the font installed in the machine
                PdfFont font = new PdfTrueTypeFont(fontStream, 14);
                //Draw the text.
                graphics.DrawString("Hello World!!!", font, PdfBrushes.Black, new PointF(0, 0));
                //Saving the PDF to the MemoryStream.
                MemoryStream stream = new MemoryStream();
                document.Save(stream);
                //Set the position as '0'.
                stream.Position = 0;
                //Download the PDF document in the browser.
                FileStreamResult fileStreamResult = new FileStreamResult(stream, "application/pdf");
                fileStreamResult.FileDownloadName = "Sample.pdf";
                return fileStreamResult;
            }
            catch (Exception ex)
            {
                ViewBag.Message = "Message : " + ex.Message.ToString() + "StackTrace : " + ex.StackTrace.ToString();
            }
            return View("Index");
        }  
    }
}
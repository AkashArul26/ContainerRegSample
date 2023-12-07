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
using Microsoft.AspNetCore.Http;

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
            string errorMessage = string.Empty;
            if (button == null)
                return View("Index");

            if (Request.Form.Files != null)
            {
                if (Request.Form.Files.Count == 0)
                {
                    ViewBag.Message = string.Format("Browse a Word document and then click the button to convert as a PDF document");
                    return View("Index");
                }
                // Gets the extension from file.
                string extension = Path.GetExtension(Request.Form.Files[0].FileName).ToLower();
                // Compares extension with supported extensions.
                if (extension == ".ttf")
                {
                    MemoryStream fontStream = new MemoryStream();
                    Request.Form.Files[0].CopyTo(fontStream);
                    try
                    {
                        //Create a new PDF document.
                        PdfDocument document = new PdfDocument();
                        //Add a page to the document.
                        PdfPage page = document.Pages.Add();

                        //Create PDF graphics for the page.
                        PdfGraphics graphics = page.Graphics;
                        //Use the font installed in the machine
                        PdfFont font = new PdfTrueTypeFont(fontStream, 14);
                        //Draw the text.
                        graphics.DrawString("Hello World!!!", font, PdfBrushes.Black, new PointF(0, 0));
                        //Saving the PDF to the MemoryStream.
                        MemoryStream pdfStream = new MemoryStream();
                        document.Save(pdfStream);
                        //Set the position as '0'.
                        pdfStream.Position = 0;
                        return File(pdfStream, "application/pdf", "WordToPDF.pdf");
                        ////Download the PDF document in the browser.
                        //FileStreamResult fileStreamResult = new FileStreamResult(stream, "application/pdf");
                        //fileStreamResult.FileDownloadName = "Sample.pdf";
                        //return fileStreamResult;
                        ////Open using Syncfusion
                        //using (WordDocument document = new WordDocument(stream, FormatType.Docx))
                        //{
                        //    stream.Dispose();
                        //    // Creates a new instance of DocIORenderer class.
                        //    using (DocIORenderer render = new DocIORenderer())
                        //    {
                        //        // Converts Word document into PDF document
                        //        using (PdfDocument pdf = render.ConvertToPDF(document))
                        //        {
                        //            MemoryStream memoryStream = new MemoryStream();
                        //            // Save the PDF document
                        //            pdf.Save(memoryStream);
                        //            memoryStream.Position = 0;
                        //            return File(memoryStream, "application/pdf", "WordToPDF.pdf");
                        //        }
                        //    }
                        //}
                    }
                    catch (Exception ex)
                    {
                        ViewBag.Message = "Message : " + ex.Message.ToString() + "StackTrace : " + ex.StackTrace.ToString();
                    }
                }
                else
                {
                    ViewBag.Message = string.Format("Please choose Word format document to convert to PDF");
                }
            }
            else
            {
                ViewBag.Message = string.Format("Browse a Word document and then click the button to convert as a PDF document");
            }
            return View("Index");
            try
            {
                Assembly assem = typeof(HomeController).Assembly;

                //string name = assem.GetName().Name;
                //throw new Exception(name);

                string fontstring = assem.GetName().Name + "\\arial.ttf";
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
using Microsoft.AspNetCore.Mvc;
using System.IO;
using System.Diagnostics;
using WordToPDFWebApp.Models;
using Microsoft.Extensions.Hosting.Internal;
using Microsoft.AspNetCore.Hosting.Server;
using System.Reflection;
using Microsoft.AspNetCore.Http;
using static System.Net.Mime.MediaTypeNames;
using SkiaSharp;
using System.Xml.Linq;
using System.Reflection.Metadata;

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
                        fontStream.Position = 0;
                        SKTypeface tfstream = SKTypeface.FromStream(fontStream);
                        throw new Exception(tfstream.FamilyName);
                        SKImageInfo imageInfo = new SKImageInfo(300, 250);
                        using (SKSurface surface = SKSurface.Create(imageInfo))
                        {
                            SKCanvas canvas = surface.Canvas;
                            canvas.Clear(SKColors.White);
                            using (SKPaint paint = new SKPaint())
                            using (SKTypeface tf = SKTypeface.FromStream(fontStream))
                            {
                                paint.Color = SKColors.Black;
                                paint.IsAntialias = true;
                                paint.TextSize = 48;
                                canvas.DrawText("Hello world", 50, 50, paint);
                            }
                            using (SKImage image = surface.Snapshot())
                            using (SKData data = image.Encode(SKEncodedImageFormat.Png, 100))
                            {
                                MemoryStream memoryStream = new MemoryStream(data.ToArray());
                                memoryStream.Position = 0;
                                return File(memoryStream, "application/png", "image.png");
                            }
                        }
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
        }
    }
}
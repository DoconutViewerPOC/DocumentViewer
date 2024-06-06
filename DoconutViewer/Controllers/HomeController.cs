using DoconutViewer.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Caching.Memory;
using System.Diagnostics;
using System.Text;

/* Please ensure that the three DLLs from .NET Core Setup are referenced or copied to the bin folder */
using Doconut.Configs;
using Doconut.Configs.View;

using Doconut; //referenced from DLLs

namespace DoconutViewer.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IWebHostEnvironment _hostingEnvironment;
        private readonly IMemoryCache _cache;
        private readonly IHttpContextAccessor _accessor;

        public HomeController(ILogger<HomeController> logger, IWebHostEnvironment hostingEnvironment, IMemoryCache cache, IHttpContextAccessor httpContextAccessor)
        {
            _logger = logger;
            _hostingEnvironment = hostingEnvironment;
            _cache = cache;
            _accessor = httpContextAccessor;
        }

        [HttpGet]
        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public IActionResult OpenDocument(string fileName)
        {
            var pathToFile = Path.Combine(Path.Combine(_hostingEnvironment.WebRootPath, "files"), fileName);

            if (!System.IO.File.Exists(pathToFile))
            {
                Response.StatusCode = 404;
                return Content($"File does not exists: {pathToFile}");
            }

            var fileInfo = new FileInfo(pathToFile);

            var licenseFilePath = Path.Combine(_hostingEnvironment.WebRootPath, "Doconut.lic");

            var docViewer = new Viewer(_cache, _accessor, licenseFilePath);

            var documentOptions = new DocOptions
            {
                Password = "",
                ImageResolution = 200,
                Watermark = "^Sample Copy~Red~24~Verdana~50~-45",
                TimeOut = 30
            };

            BaseConfig? config = null;

            switch (fileInfo.Extension.ToUpper())
            {
                case ".DOC":
                case ".DOCX":
                case ".ODT":
                case ".RTF":
                    config = new WordConfig { ConvertPdf = true, PdfConfig = new PdfConfig { ExtractHyperlinks = true, HyperlinksPageCount = 5 } };
                    break;
                case ".XLS":
                case ".XLSX":
                case ".ODS":
                case ".CSV":
                    config = new ExcelConfig { SplitWorksheets = true, ShowEmptyWorkSheets = false, DocumentCulture = "en-US", PaperSize = ExcelPaperSize.PaperA3, AutoFitContents = true };
                    break;
                case ".PPT":
                case ".PPTX":
                case ".ODP":
                    config = new PptConfig();
                    break;
                case ".DWG":
                case ".DXF":
                case ".DGN":
                    config = new CadConfig { DefaultRender = true, ShowColor = false, WhiteBackground = true, ShowModel = true, ShowLayouts = true, LineWidth = 1, Check3DSolid = false };
                    break;
                case ".EML":
                case ".MSG":
                    config = new EmailConfig { EmailEncoding = Encoding.UTF8 };
                    break;
                case ".PDF":
                    config = new PdfConfig { ExtractHyperlinks = true, HyperlinksPageCount = 5 };
                    break;
                case ".BMP":
                case ".CDR":
                case ".CMX":
                case ".DCM":
                case ".DNG":
                case ".EPS":
                case ".GIF":
                case ".ICO":
                case ".JPG":
                case ".JPEG":
                case ".PNG":
                case ".PSD":
                case ".TGA":
                case ".WEBP":
                    config = new ImageConfig { MaxImagePixelSize = 2000, TransparentPng = false };
                    break;
                case ".TXT":
                    config = new WordConfig { PaperSize = DocPaperSize.A4 };
                    break;
                case ".MPP":
                case ".MPPX":
                    config = new ProjectConfig { ExportPdfA = true, PaperSize = MppPaperSize.A3 };
                    break;
                case ".VSD":
                case ".VSDX":
                    config = new VisioConfig { ExportPdfA = true };
                    break;
            }


            try
            {
                var token = docViewer.OpenDocument(pathToFile, config, documentOptions);

                return Content(token);
            }
            catch (Exception e)
            {
                Response.StatusCode = 500;
                return Content(e.Message);
            }

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
    }
}

using Microsoft.AspNetCore.Cors;
using Microsoft.AspNetCore.Mvc;
using System.Threading.Tasks;
using Report.Models;
using Report.Helpers;
using Microsoft.Office.Interop.Word;

namespace Report.Controllers {
    [Route("api/createreport")]
    [ApiController]
    public class ReportController : Controller {
        public string config(string name) {
            return AppSettings.Configuration.GetSection(name).Value;
        }

        [EnableCors("_myAllowSpecificOrigins")]
        [HttpPost]
        public SaveFileResult ReplaceOpenXML([FromBody] dynamic content) {
            var pdf = new PDF();
            var serverAddr = $"{this.Request.Scheme}://{this.Request.Host}{this.Request.PathBase}/";
            var targetPdfName = pdf.CreatePDFReport(content, serverAddr);
            var returnResponse = new SaveFileResult();
            returnResponse.FileName = targetPdfName;
            ReportUtils reportUtils = new ReportUtils();
            reportUtils.DeleteOldFiles();

            return returnResponse;
        }
    }

    [Route("api/test")]
    [ApiController]
    public class ReportController1 : Controller {
        public string config(string name) {
            return AppSettings.Configuration.GetSection(name).Value;
        }

        [EnableCors("_myAllowSpecificOrigins")]
        [HttpGet]
        public string Test() {
            Application app = new Microsoft.Office.Interop.Word.Application();
            return "HELLO";
        }
    }
}
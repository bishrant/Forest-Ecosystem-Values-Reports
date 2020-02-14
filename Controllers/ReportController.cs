using Microsoft.AspNetCore.Cors;
using Microsoft.AspNetCore.Mvc;
using System.Threading.Tasks;
using Report.Models;
using Report.Helpers;

namespace Report.Controllers {
    [Route("api/createreport")]
    [ApiController]
    public class ReportController : Controller {
        public string config(string name) {
            return AppSettings.Configuration.GetSection(name).Value;
        }

        [EnableCors("AllowOrigin")]
        [HttpPost]
        public async Task<SaveFileResult> ReplaceOpenXML([FromBody] dynamic content) {
            var pdf = new PDF();
            var targetPdfName = pdf.CreatePDFReport();
            var returnResponse = new SaveFileResult();
            returnResponse.FileName = targetPdfName;
            return returnResponse;
        }
    }
}
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;

namespace Report.Helpers
{
    public class PDF
    {
        public string config(string name) {
            return AppSettings.Configuration.GetSection(name).Value;
        }

        public string CreatePDFReport(dynamic reportData, string serverAddr) {
            var jsonStr = reportData.GetValue("stats").ToString();
            string mapURL = reportData.GetValue("mapURL").ToString();
            ReportData data = ReportData.FromJson(jsonStr);
            string directory = Directory.GetCurrentDirectory();
            var setting = new OpenSettings();
            setting.AutoSave = false;
            var dateUtils = new DateUtils();
            StringUtils stringUtils = new StringUtils();
            ReportUtils ReportUtils = new ReportUtils();
            string fileName = config("filePrefix")+ "_" + DateTime.Now.ToString("MM-dd-yyyy") + "_" + DateTime.Now.ToString("HHmmss");
            string tempFileName = directory + "\\wwwroot\\reports\\" + fileName + ".docx";
            string templateFileName = directory + "\\wwwroot\\template\\"+ config("template");
            try  {
                File.Copy(templateFileName, tempFileName);
            } catch (Exception e) { throw(e); }

            ReportUtils.ReplaceImage(tempFileName, mapURL, config("mapImageId"));

            SummaryReportResults summary = data.SummaryReportResults;

            // get the individual summary objects for replacing to the openxml later
            SummaryResult[] summaryArray = data.SummaryResults;
            SummaryResult airquality = summaryArray.Where(s => s.EcosystemService == "airquality").ToArray()[0];
            SummaryResult biodiversity = summaryArray.Where(s => s.EcosystemService == "biodiversity").ToArray()[0];
            SummaryResult carbon = summaryArray.Where(s => s.EcosystemService == "carbon").ToArray()[0];
            SummaryResult cultural = summaryArray.Where(s => s.EcosystemService == "cultural").ToArray()[0];
            SummaryResult watershed = summaryArray.Where(s => s.EcosystemService == "watershed").ToArray()[0];
            SummaryResult total = summaryArray.Where(s => s.EcosystemService == "total").ToArray()[0];

            //Create report params dictionary
            Dictionary<string, string> reportParamsDict = new Dictionary<string, string>();
            reportParamsDict.Add("ZZAoiForestAcresZZ", summary.AoiForestAcres);
            reportParamsDict.Add("ZZAoiUrbanAcresZZ", ""+summary.AoiUrbanAcres);
            reportParamsDict.Add("ZZAoiRuralAcresZZ", summary.AoiRuralAcres);

            reportParamsDict.Add("ZZairqualityAvgZZ", airquality.ForestAverageValue);
            reportParamsDict.Add("ZZairqualityTotalZZ", airquality.ForestTotalValue);

            reportParamsDict.Add("ZZbiodiversityAvgZZ", biodiversity.ForestAverageValue);
            reportParamsDict.Add("ZZbiodiversityTotalZZ", biodiversity.ForestTotalValue);

            reportParamsDict.Add("ZZcarbonAvgZZ", carbon.ForestAverageValue);
            reportParamsDict.Add("ZZcarbonTotalZZ", carbon.ForestTotalValue);

            reportParamsDict.Add("ZZculturalAvgZZ", cultural.ForestAverageValue);
            reportParamsDict.Add("ZZculturalTotalZZ", cultural.ForestTotalValue);

            reportParamsDict.Add("ZZwatershedAvgZZ", watershed.ForestAverageValue);
            reportParamsDict.Add("ZZwatershedTotalZZ", watershed.ForestTotalValue);

            reportParamsDict.Add("ZZtotalAvgZZ", total.ForestAverageValue);
            reportParamsDict.Add("ZZtotalTotalZZ", total.ForestTotalValue);

            string[] ecosystems = new string[] { "airquality", "biodiversity", "carbon", "cultural", "watershed", "total"};
            foreach (string eco in ecosystems) {
                reportParamsDict.Add("ZZ"+eco+"UrbanZZ", (string)summary[eco + "_urbanThousandDollarsPerYear"]);
                reportParamsDict.Add("ZZ" + eco + "RuralZZ", (string)summary[eco + "_ruralThousandDollarsPerYear"]);
                reportParamsDict.Add("ZZ" + eco + "Total1ZZ", (string)summary[eco + "_totalThousandDollarsPerYear"]);
            }

            stringUtils.SearchAndReplace(tempFileName, reportParamsDict);
            DocxToPDF PDFConvertor = new DocxToPDF();
            PDFConvertor.ConvertDocxToPDF(tempFileName);
            string pdfPath = serverAddr+ "reports/" + fileName + ".pdf";;
            File.Delete(tempFileName);
            return pdfPath;
        }
    }
}

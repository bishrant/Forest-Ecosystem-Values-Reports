using DocumentFormat.OpenXml.Packaging;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Report.Helpers;

namespace Report.Helpers
{
    public class PDF
    {
        public string config(string name) {
            return AppSettings.Configuration.GetSection(name).Value;
        }

        public void ConvertDocxToPDF(string docPath) {
            Application app = new Application();
            string pdfPath = docPath.Replace(".docx", ".pdf");
            Document wordDoc = null;
            try {
                wordDoc = app.Documents.Open(docPath);
                wordDoc.ExportAsFixedFormat(pdfPath, WdExportFormat.wdExportFormatPDF);
                wordDoc.Close();
                app.Quit();
                wordDoc = null;
                app = null;
            } catch (Exception ex) {
                throw (ex);
            } finally {
                if (wordDoc != null) { wordDoc.Close(); wordDoc = null; }
                if (app != null) { app.Quit(); app = null; }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        public static void SearchAndReplace(string document, Dictionary<string, string> myDictionary) {

            string docText;
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true)) {
                docText = null;
                using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream())) {
                    docText = sr.ReadToEnd();
                }

                foreach (KeyValuePair<string, string> entry in myDictionary) {
                    Regex regexText = new Regex(entry.Key);
                    var paramValue = entry.Value;
                    docText = regexText.Replace(docText, paramValue);
                }

                using (StreamWriter sw = new StreamWriter(
                          wordDoc.MainDocumentPart.GetStream(FileMode.Create))) {
                    sw.Write(docText);
                }
                wordDoc.Close();
            }
        }

        public string CreatePDFReport(dynamic reportData)
        {
            var jsonStr = reportData.GetProperty("stats").GetString();
            string mapURL = reportData.GetProperty("mapURL").GetString();
            ReportData data = ReportData.FromJson(jsonStr);
            string directory = Directory.GetCurrentDirectory();
            var setting = new OpenSettings();
            setting.AutoSave = false;
            var dateUtils = new DateUtils();
            var stringUtils = new StringUtils();
            ReportUtils ReportUtils = new ReportUtils();
            string fileName = config("filePrefix")+ "_" + DateTime.Now.ToString("MM-dd-yyyy") + "_" + DateTime.Now.ToString("HHmmss") + ".docx";
            string tempFileName = directory + "\\template\\" + fileName;
            string templateFileName = directory + "\\template\\"+ config("template");
            try  {
                File.Copy(templateFileName, tempFileName);
            } catch (Exception e) { Console.WriteLine(e); }

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

            string[] ecosystems = new string[] { "airquality", "biodiversity", "carbon", "cultural", "watershed", "total" };
            foreach (string eco in ecosystems) {
                reportParamsDict.Add("ZZ"+eco+"UrbanZZ", (string)summary[eco + "_urbanThousandDollarsPerYear"]);
                reportParamsDict.Add("ZZ" + eco + "RuralZZ", (string)summary[eco + "_ruralThousandDollarsPerYear"]);
                reportParamsDict.Add("ZZ" + eco + "Total1ZZ", (string)summary[eco + "_totalThousandDollarsPerYear"]);
            }

            SearchAndReplace(tempFileName, reportParamsDict);

            ConvertDocxToPDF(tempFileName);
            string pdfPath = tempFileName.Replace(".docx", ".pdf");

            ////ImagePart imagePartLandOwner;
            //using (WordprocessingDocument doc = WordprocessingDocument.Open(tempFileName, true)) {
            //    string docText = null;
            //    using (StreamReader sr = new StreamReader(doc.MainDocumentPart.GetStream())) {
            //        docText = sr.ReadToEnd();
            //    }

            //    TextInfo myTI = new CultureInfo("en-US", false).TextInfo;
            //    SummaryReportResults summary = data.SummaryReportResults;
            //    IDictionary<string, string> map = new Dictionary<string, string>()  {
            //       {"ZZAoiForestAcresZZ", "1000"},
            //       {"ZZaoiRuralAcresZZ", summary.AoiRuralAcres}
            //    };

            //    var regex = new Regex(String.Join("|", map.Keys.Select(k => Regex.Escape(k))));
            //    docText = regex.Replace(docText, m => stringUtils.escape(map[m.Value]));

            //    using (StreamWriter sw = new StreamWriter(doc.MainDocumentPart.GetStream(FileMode.Create))) {
            //        sw.Write(docText);
            //    }
            //    //ImagePart imagePart = (ImagePart)doc.MainDocumentPart.GetPartById("rId7");

            //    //var webClient = new WebClient();
            //    //var signatureUrl = "https://www.nuget.org/Content/gallery/img/default-package-icon.svg";
            //    //byte[] imageBytes = webClient.DownloadData(signatureUrl);
            //    //BinaryWriter writer = new BinaryWriter(imagePart.GetStream());
            //    //writer.Write(imageBytes);
            //    //writer.Close();
            //    //imagePartLandOwner = (ImagePart)doc.MainDocumentPart.GetPartById("rId9");
            //    doc.Save();
            //};
            //string targetPdfName = stringUtils.RemoveFilenameInvalidChars("UAS Flight Authorization Application_" + 
            //    stringUtils.cleanTxt(content.feature.attributes["pilotName"]) +
            //    "_" + missionDate.Date.ToString("MM-dd-yyyy") + "_" + DateTime.Now.ToString("HHmmss") + ".pdf");
            //string targetPdfPath = directory + "\\wwwroot\\UASApplications\\" + targetPdfName;
            //Application app = new Application();
            //Document worddoc = app.Documents.Open(tempFileName);
            //worddoc.ExportAsFixedFormat(targetPdfPath, WdExportFormat.wdExportFormatPDF);
            //worddoc.Close();
            //app.Quit();
            //File.Delete(tempFileName);
            return pdfPath;
        }
    }
}

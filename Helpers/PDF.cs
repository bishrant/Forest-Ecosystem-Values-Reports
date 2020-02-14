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

namespace Report.Helpers
{
    public class PDF
    {
        public string config(string name) {
            return AppSettings.Configuration.GetSection(name).Value;
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

        public string CreatePDFReport()
        {
            ReportData data = ReportData.FromJson("{\"summaryReportResults\":{\"airquality_forestAnnual\":\"0\",\"airquality_forestValuePerAcre\":\"0.0\",\"airquality_ruralThousandDollarsPerYear\":\"-\",\"airquality_totalThousandDollarsPerYear\":\"0.0\",\"airquality_urbanThousandDollarsPerYear\":\"0.0000\",\"aoiForestAcres\":\"15,742\",\"aoiRuralAcres\":\"15,742\",\"aoiUrbanAcres\":\"0\",\"biodiversity_forestAnnual\":\"3.7 million\",\"biodiversity_forestValuePerAcre\":\"232.0\",\"biodiversity_ruralThousandDollarsPerYear\":\"3,652.2\",\"biodiversity_totalThousandDollarsPerYear\":\"3,652.2\",\"biodiversity_urbanThousandDollarsPerYear\":\"0.0000\",\"carbon_forestAnnual\":\"1.4 million\",\"carbon_forestValuePerAcre\":\"88.8\",\"carbon_ruralThousandDollarsPerYear\":\"1,397.4\",\"carbon_totalThousandDollarsPerYear\":\"1,397.4\",\"carbon_urbanThousandDollarsPerYear\":\"0.0000\",\"cultural_forestAnnual\":\"14.2 million\",\"cultural_forestValuePerAcre\":\"899.8\",\"cultural_ruralThousandDollarsPerYear\":\"14,164.8\",\"cultural_totalThousandDollarsPerYear\":\"14,164.8\",\"cultural_urbanThousandDollarsPerYear\":\"0.0000\",\"total_forestAnnual\":\"23.2 million\",\"total_forestValuePerAcre\":\"1,476.1\",\"total_ruralThousandDollarsPerYear\":\"23,236.8\",\"total_totalThousandDollarsPerYear\":\"23,236.8\",\"total_urbanThousandDollarsPerYear\":\"0.0000\",\"watershed_forestAnnual\":\"4.0 million\",\"watershed_forestValuePerAcre\":\"255.5\",\"watershed_ruralThousandDollarsPerYear\":\"4,022.4\",\"watershed_totalThousandDollarsPerYear\":\"4,022.4\",\"watershed_urbanThousandDollarsPerYear\":\"0.0000\"},\"summaryResults\":[{\"AverageValueUnits\":\"$/acre/year\",\"TotalValueUnits\":\"thousand $/year\",\"ecosystemService\":\"airquality\",\"forestAverageValue\":\"0.0\",\"forestTotalValue\":\"0.0\",\"ruralAverageValue\":\"-\",\"ruralTotalValue\":\"-\",\"urbanAverageValue\":\"0.0\",\"urbanTotalValue\":\"0.0000\"},{\"AverageValueUnits\":\"$/acre/year\",\"TotalValueUnits\":\"thousand $/year\",\"ecosystemService\":\"biodiversity\",\"forestAverageValue\":\"232.0\",\"forestTotalValue\":\"3,652.2\",\"ruralAverageValue\":\"232.0\",\"ruralTotalValue\":\"3,652.2\",\"urbanAverageValue\":\"0.0\",\"urbanTotalValue\":\"0.0000\"},{\"AverageValueUnits\":\"$/acre/year\",\"TotalValueUnits\":\"thousand $/year\",\"ecosystemService\":\"carbon\",\"forestAverageValue\":\"88.8\",\"forestTotalValue\":\"1,397.4\",\"ruralAverageValue\":\"88.8\",\"ruralTotalValue\":\"1,397.4\",\"urbanAverageValue\":\"0.0\",\"urbanTotalValue\":\"0.0000\"},{\"AverageValueUnits\":\"$/acre/year\",\"TotalValueUnits\":\"thousand $/year\",\"ecosystemService\":\"cultural\",\"forestAverageValue\":\"899.8\",\"forestTotalValue\":\"14,164.8\",\"ruralAverageValue\":\"899.8\",\"ruralTotalValue\":\"14,164.8\",\"urbanAverageValue\":\"0.0\",\"urbanTotalValue\":\"0.0000\"},{\"AverageValueUnits\":\"$/acre/year\",\"TotalValueUnits\":\"thousand $/year\",\"ecosystemService\":\"watershed\",\"forestAverageValue\":\"255.5\",\"forestTotalValue\":\"4,022.4\",\"ruralAverageValue\":\"255.5\",\"ruralTotalValue\":\"4,022.4\",\"urbanAverageValue\":\"0.0\",\"urbanTotalValue\":\"0.0000\"},{\"AverageValueUnits\":\"$/acre/year\",\"TotalValueUnits\":\"thousand $/year\",\"ecosystemService\":\"total\",\"forestAverageValue\":\"1,476.1\",\"forestTotalValue\":\"23,236.8\",\"ruralAverageValue\":\"1,476.1\",\"ruralTotalValue\":\"23,236.8\",\"urbanAverageValue\":\"0.0\",\"urbanTotalValue\":\"0.0000\"}]}");
            string directory = Directory.GetCurrentDirectory();
            var setting = new OpenSettings();
            setting.AutoSave = false;
            var dateUtils = new DateUtils();
            var stringUtils = new StringUtils();
            string fileName = config("filePrefix")+ "_" + DateTime.Now.ToString("MM-dd-yyyy") + "_" + DateTime.Now.ToString("HHmmss") + ".docx";
            string tempFileName = directory + "\\template\\" + fileName;
            string templateFileName = directory + "\\template\\"+ config("template");
            try  {
                File.Copy(templateFileName, tempFileName);
            } catch (Exception e) { Console.WriteLine(e); }

            SummaryReportResults summary = data.SummaryReportResults;

            //Create report params dictionary
            Dictionary<string, string> reportParamsDict = new Dictionary<string, string>();
            reportParamsDict.Add("ZZAoiForestAcresZZ", summary.AoiForestAcres);
            reportParamsDict.Add("ZZAoiUrbanAcresZZ", ""+summary.AoiUrbanAcres);
            reportParamsDict.Add("ZZAoiRuralAcresZZ", summary.AoiRuralAcres);

            SearchAndReplace(tempFileName, reportParamsDict);

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
            return tempFileName;
        }
    }
}

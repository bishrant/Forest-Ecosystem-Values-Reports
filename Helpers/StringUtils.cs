using System;
using System.Text.RegularExpressions;
using System.IO;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;

namespace Report.Helpers
{
    public class StringUtils
    {
        public void SearchAndReplace(string document, Dictionary<string, string> myDictionary) {
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
        public string convetToString(object o) {
            string val = "";
            if (Convert.ToString(o) != "") {
                val = Convert.ToString(o);
            }
            return val;
        }

        public string escape(dynamic text) {
            var t = Convert.ToString(text);
            if (t == "" || t == null) {
                return t;
            } else {
                t = Regex.Replace(t, "&", "&amp;");
                t = Regex.Replace(t, "'", "&apos;");
                t = Regex.Replace(t, "\"", "&quot;");
                t = Regex.Replace(t, ">", "&gt;");
                t = Regex.Replace(t, "<", "&lt;");
                return t;
            }
        }

        public string cleanTxt(dynamic text) {
            return Regex.Replace(Convert.ToString(text), @"(\s+|@|&|'|\(|\)|<|>|#)", "");
        }

        public string RemoveFilenameInvalidChars(string filename) {
            return string.Concat(filename.Split(Path.GetInvalidFileNameChars()));
        }

    }
}

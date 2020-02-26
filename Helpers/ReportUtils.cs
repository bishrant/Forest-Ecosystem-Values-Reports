using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;

using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Report.Helpers
{
    public class ReportUtils
    {

        public void DeleteOldFolders( string _newDocumentsFolderPath)
        {
            string[] folders = Directory.GetDirectories(_newDocumentsFolderPath);
            foreach (string folder in folders)
            {
                DirectoryInfo di = new DirectoryInfo(folder);
                if (di.LastAccessTime < DateTime.Now.AddDays(-1))
                {
                    string[] files = Directory.GetFiles(folder);
                    foreach (string file in files)
                    {
                        FileInfo fi = new FileInfo(file);
                        fi.Delete();
                    }
                    di.Delete();
                }
            }
        }

        // To search and replace content in a document part.
        public void SearchAndReplace(string document, Dictionary<string, string> myDictionary)
        {

            string docText;
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
            {
                docText = null;
                using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }

                foreach (KeyValuePair<string, string> entry in myDictionary)
                {
                    Regex regexText = new Regex(entry.Key);
                    var paramValue = entry.Value;
                    docText = regexText.Replace(docText, paramValue);
                }

                using (StreamWriter sw = new StreamWriter(
                          wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                {
                    sw.Write(docText);
                }
                wordDoc.Close();
            }
        }

        public void RemoveSections(string destinationFile, List<int> sectionToRemoveList) {
            using (WordprocessingDocument myDoc = WordprocessingDocument.Open(destinationFile, true)) {
                MainDocumentPart mainPart = myDoc.MainDocumentPart;
                List<ParagraphProperties> paraProps = mainPart.Document.Descendants<ParagraphProperties>()
                .Where(pPr => IsSectionProps(pPr)).ToList();

                int pCounter = 0;
                foreach (ParagraphProperties pPr in paraProps) {
                    pCounter++;
                    if (sectionToRemoveList.Contains(pCounter)) {
                        var p = pPr.Parent.PreviousSibling();
                        while (p != null && !IsSectionProps(p.GetFirstChild<ParagraphProperties>())) {
                            var p1 = p.PreviousSibling();
                            p.Remove();
                            p = p1;
                        }
                        pPr.Parent.Remove();
                    }
                }

                mainPart.Document.Save();
            }
        }

        public void ReplaceImage(string destinationFile, string mapImageUrl, string imageId) {
            WordprocessingDocument m_wordProcessingDocument = WordprocessingDocument.Open(destinationFile, true);
            MainDocumentPart m_mainDocPart = m_wordProcessingDocument.MainDocumentPart;

            // go through the document and pull out the inline image elements
            IEnumerable<Drawing> imageElements = from run in m_mainDocPart.Document.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>()
                                                 where run.Descendants<Drawing>().First() != null
                                                 select run.Descendants<Drawing>().First();

            ImagePart imagePart = (ImagePart)m_mainDocPart.GetPartById(imageId);

            var webClient = new WebClient();
            byte[] m_imageInBytes = webClient.DownloadData(mapImageUrl);
            BinaryWriter writer = new BinaryWriter(imagePart.GetStream());
            writer.Write(m_imageInBytes);
            writer.Close();

            m_wordProcessingDocument.Close();
        }

        private bool IsSectionProps(ParagraphProperties pPr) {
            if (pPr != null) {
                SectionProperties sectPr = pPr.GetFirstChild<SectionProperties>();
                if (sectPr == null)
                    return false;
                else
                    return true;
            }
            return false;
        }
    }


}

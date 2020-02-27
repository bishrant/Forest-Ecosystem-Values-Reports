using Microsoft.Office.Interop.Word;
using System;

namespace Report.Helpers {
    public class DocxToPDF {
        public void ConvertDocxToPDF(string docPath) {
            Application app = new Microsoft.Office.Interop.Word.Application();
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
    }
}

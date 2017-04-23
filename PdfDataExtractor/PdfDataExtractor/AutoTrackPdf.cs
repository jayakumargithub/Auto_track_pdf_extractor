using System.Configuration;
using System.IO;
using System.Linq;
using Bytescout.PDFExtractor;
using log4net;

namespace PdfDataExtractor
{
    public class AutoTrackPdf
    {
        private static readonly ILog Logger =
            LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public void Extract()
        {
            string xlsFile = string.Empty;
            string _pdfFile = string.Empty;

            var rootFolder = FileLocation.CreateBaseFolderIsfNotExists();

            var files = Directory.GetFiles(rootFolder + ConfigurationManager.AppSettings["PdfFileLocation"]).Where(x => x.EndsWith(".pdf")).ToArray();
            if (files.Length == 0)
            {
                Logger.Info("No Pdf File Found");
                return;
            }


            var extractor = new XLSExtractor();
            FileInfo info = null;
            extractor.RegistrationName = "NOT - FOR - RESALE - SINGLE - ENDUSER - ONLY - NO - PRIVATE - SUPPORT - jayakumar.m.h@gmail.com";
            extractor.RegistrationKey = "10C9-AC43-B36E-997C-CCCF-BDEB-C9D";

            _pdfFile = files[0];

            Logger.Info("Total file Loaded: " + files.Length);
            Logger.Info("Pdf file loaded: " + _pdfFile);
            extractor.PageDataCaching = PageDataCaching.None;

            Logger.Info("Xsl file loaded: " + xlsFile);
            extractor.AutoAlignColumnsToHeader = true;
            Logger.Info("Pdf file extraction started");
            Logger.Info("Loaded file:" + files[0]);
            info = new FileInfo(_pdfFile);
            xlsFile = FileLocation.GetFilePath(FileLocationEnum.Xls);
            xlsFile += info.Name.Split('.')[0] + ".xls";

            try
            {
                extractor.LoadDocumentFromFile(files[0]);
                extractor.SaveToXLSFile(xlsFile);
                extractor.Reset();
            }
            catch (PDFExtractorException ex)
            {
                Logger.Info("PDFExtractorException: " + ex.Message);
            }
            if (extractor.IsDocumentLoaded)
            {
                extractor.Dispose();
            }



            XlsReaderService.ReadXls(xlsFile);
            try
            {
                var destination = FileLocation.GetFilePath(FileLocationEnum.ProcessedPdf);
                var fileName = info.Name.Split('.')[0];
                info.MoveTo(destination + fileName + ".pdf");

            }
            catch (IOException ex)
            {
                Logger.Info("Delete Error: " + ex.Message);
            }
        }


    }
}
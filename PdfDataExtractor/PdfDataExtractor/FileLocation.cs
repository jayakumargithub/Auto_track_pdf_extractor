using System.Configuration;
using System.IO;
using log4net;

namespace PdfDataExtractor
{
    public class FileLocation
    {
        private static readonly ILog Logger =
            LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public static string GetFilePath(FileLocationEnum location)
        {
            var loc = string.Empty;
            var rootFolder = ConfigurationManager.AppSettings["RootFolder"];

            try
            {
                switch (location)
                {
                    case FileLocationEnum.Pdf:
                        loc = rootFolder + ConfigurationManager.AppSettings["PdfFileLocation"];
                        break;
                    case FileLocationEnum.CN:
                        loc = rootFolder + ConfigurationManager.AppSettings["txtFolder"] + "\\CN" +
                              ".txt" +
                              "";
                        break;
                    case FileLocationEnum.ST:
                        loc = rootFolder + ConfigurationManager.AppSettings["txtFolder"] + "\\ST.txt";
                        break;
                    case FileLocationEnum.FT:
                        loc = rootFolder + ConfigurationManager.AppSettings["txtFolder"] + "\\FT.txt";
                        break;
                    case FileLocationEnum.Name:
                        loc = rootFolder + ConfigurationManager.AppSettings["txtFolder"] + "\\Name.txt";
                        break;
                    case FileLocationEnum.Xls:
                        loc = rootFolder + ConfigurationManager.AppSettings["XlsFileLocation"];
                        break;
                    case FileLocationEnum.ProcessedPdf:
                        loc = rootFolder + ConfigurationManager.AppSettings["ProcessedPdf"];
                        break;
                    default:
                        break;

                }
            }
            catch (ConfigurationException ex)
            {
                Logger.Info("FileLocation error:" + ex.Message);
            }

            return loc;
        }


        public static string CreateBaseFolderIsfNotExists()
        {

            var rootFolder = ConfigurationManager.AppSettings["RootFolder"];
            if (!Directory.Exists(rootFolder))
            {
                Directory.CreateDirectory(rootFolder);
            }


            var pdfDropFolder = ConfigurationManager.AppSettings["PdfFileLocation"];
            var txtFolder = ConfigurationManager.AppSettings["txtFolder"];
            var processedPdfFolder = ConfigurationManager.AppSettings["processedPdf"];
            var xlsFileLocation = ConfigurationManager.AppSettings["XlsFileLocation"];


            if (!Directory.Exists(rootFolder + pdfDropFolder))
            {
                Directory.CreateDirectory(rootFolder + pdfDropFolder);
            }
            if (!Directory.Exists(rootFolder + txtFolder))
            {
                Directory.CreateDirectory(rootFolder + txtFolder);
            }
            if (!Directory.Exists(rootFolder + processedPdfFolder))
            {
                Directory.CreateDirectory(rootFolder + processedPdfFolder);
            }
            if (!Directory.Exists(rootFolder + xlsFileLocation))
            {
                Directory.CreateDirectory(rootFolder + xlsFileLocation);
            }
            return rootFolder;
        }
    }
}
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using log4net;

namespace PdfDataExtractor
{
    public static class FileWriter
    {
        private static readonly ILog Logger =
            LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public static void WriteFile(FileType fileType, List<RacerDetail> racer)
        {
            log4net.Config.XmlConfigurator.Configure();
            var fileLoction = FileLocation.GetFilePath(FileLocationEnum.Txt);
            switch (fileType)
            {
                case FileType.CN:
                    if (File.Exists(fileLoction))
                    {
                        File.Delete(fileLoction);
                    }
                    StringBuilder sb = new StringBuilder();
                    int cnIndex = 1;
                    foreach (var item in racer)
                    {
                        sb.Append("<C" + cnIndex + "\t" + item.Number);
                        sb.Append(System.Environment.NewLine);
                        cnIndex++;
                    }
                    try
                    {
                        using (System.IO.StreamWriter file = new System.IO.StreamWriter(fileLoction, true))
                        {
                            file.Write(sb.ToString());
                        }
                    }
                    catch (UnauthorizedAccessException ex)
                    {
                        Logger.Info(ex.Message);
                    }
                    catch (Exception ex)
                    {
                        Logger.Info(ex.Message);
                    }
                    break;
                default:
                    break;
            }
            Console.ReadLine();
        }
    }
}

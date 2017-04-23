using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Text;
using log4net;

namespace PdfDataExtractor
{
    public class XlsReaderService
    {
        private static readonly ILog Logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public static void ReadXls(string xslFile)

        {
            using (OleDbConnection conn = new OleDbConnection())
            {
                DataTable dt = new DataTable();

                conn.ConnectionString = GetConnectionString(xslFile);// "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + xlsFile + ";" + "Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1;MAXSCANROWS=0'";
                try
                {
                    using (OleDbCommand comm = new OleDbCommand())
                    {
                        comm.CommandText = "Select * from [" + "Page 1" + "$]";
                        comm.Connection = conn;
                        using (OleDbDataAdapter da = new OleDbDataAdapter())
                        {
                            da.SelectCommand = comm;
                            da.Fill(dt);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logger.Info("XlsReaderService Error: " + ex.Message);
                }


                var racers = BuildTable(dt);
                WriteFile(FileType.CN, racers);
                WriteFile(FileType.ST, racers);
                WriteFile(FileType.FT, racers);
                WriteFile(FileType.Name,racers);
            }
        }



        private static string GetConnectionString(string path)
        {
            Dictionary<string, string> props = new Dictionary<string, string>();



            // XLS - Excel 2003 and Older
            //if (version == ExcelVersion.Xls)
            //{
            props["Provider"] = "Microsoft.Jet.OLEDB.4.0";
            props["Extended Properties"] = "Excel 8.0";
            props["Data Source"] = path;
            //}
            //else if (version == ExcelVersion.Xlsx)
            //{
            //    // XLSX - Excel 2007, 2010, 2012, 2013
            //    props["Provider"] = "Microsoft.ACE.OLEDB.12.0;";
            //    props["Extended Properties"] = "Excel 12.0 XML";
            //    props["Data Source"] = path;
            //}

            StringBuilder sb = new StringBuilder();

            foreach (KeyValuePair<string, string> prop in props)
            {
                sb.Append(prop.Key);
                sb.Append('=');
                sb.Append(prop.Value);
                sb.Append(';');
            }

            return sb.ToString();
        }
        public static List<RacerDetail> BuildTable(DataTable table)
        {
            List<RacerDetail> racers = new List<RacerDetail>();
            try
            {
                for (int i = 0; i < 20; i++)
                {
                    var racer = new RacerDetail();

                    DataRow row = table.Rows[i];
                    int slNo;
                    var isGreaterThanZero = int.TryParse(row.ItemArray[0].ToString(), out slNo);
                    if (!isGreaterThanZero)
                        continue;

                    racer.SlNo = row.ItemArray[0]?.ToString();
                    racer.Number = row.ItemArray[1]?.ToString();
                    racer.Name = row.ItemArray[2]?.ToString();
                    racer.Gap = row.ItemArray[3]?.ToString();
                    racer.TotalTime = row.ItemArray[4]?.ToString();
                    racer.Fastest = row.ItemArray[5]?.ToString();
                    racer.In = Convert.ToInt32(row.ItemArray[6]?.ToString());
                    racer.AverageSpeed = Convert.ToDecimal(row.ItemArray[7]?.ToString());
                    racers.Add(racer);

                }


            }
            catch (Exception ex)
            {
                Logger.Info("XlsReaderSErvce.BuildTable Error:" + ex.Message);
            }


            return racers;

        }



        public static void WriteFile(FileType fileType, List<RacerDetail> racer)
        {

            var file = FileLocationEnum.None;
            if (fileType == FileType.CN)
            {
                file = FileLocationEnum.CN;
            }
            else if (fileType == FileType.ST)
            {
                file = FileLocationEnum.ST;
            }
            else if (fileType == FileType.FT)
            {
                file = FileLocationEnum.FT;
            }
            else if (fileType == FileType.Name)
            {
                file = FileLocationEnum.Name;
            }


            try
            {
                var fileLoction = FileLocation.GetFilePath(file);
                switch (fileType)
                {
                    case FileType.CN:
                        WriteToFile(racer, fileLoction, FileType.CN);
                        break;
                    case FileType.ST:
                        WriteToFile(racer, fileLoction, FileType.ST);
                        break;
                    case FileType.FT:
                        WriteToFile(racer, fileLoction, FileType.FT);
                        break;
                    case FileType.Name:
                        WriteToFile(racer, fileLoction, FileType.Name);
                        break;
                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                Logger.Info("XlsReaderService.WriteFile Error:" + ex.Message);
            }
            Console.ReadLine();
        }

        private static void WriteToFile(List<RacerDetail> racer, string fileLoction, FileType fileType)
        {

            var racerCount = racer.Count;
            int count = 20 - racerCount;
            for (int i = 0; i <= count - 1; i++)
            {
                var rac = new RacerDetail();
                racer.Add(rac);
            }


            StringBuilder sb = new StringBuilder();
            int cnIndex = 1;
            foreach (var item in racer)
            {

                if (fileType == FileType.CN)
                {
                    sb.Append("<C" + cnIndex + "\t" + item.Number);
                }
                else if (fileType == FileType.ST)
                {
                    sb.Append("<T" + cnIndex + "\t" + item.TotalTime);
                }
                else if (fileType == FileType.Name)
                {
                    sb.Append("<T" + cnIndex + "\t" + item.Name);
                }
                else if (fileType == FileType.FT)
                {
                    if (cnIndex == 1)
                    {
                        sb.Append("<FT1" + "\t" + item.Fastest);
                    }
                }



                sb.Append(System.Environment.NewLine);
                cnIndex++;
            }
            try
            {
                using (Stream file1 = File.Open(fileLoction, System.IO.FileMode.Create))
                {
                    using (System.IO.StreamWriter file = new System.IO.StreamWriter(file1))
                    {
                        file.Write(sb);
                    }

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
        }
    }
}

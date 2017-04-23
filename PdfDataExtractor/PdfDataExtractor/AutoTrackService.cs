using System;
using System.Configuration;
using System.ServiceProcess;
using System.Timers;
using log4net;

namespace PdfDataExtractor
{
    public partial class AutoTrackService : ServiceBase
    {
        private static readonly ILog Logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private Timer timer;
        public AutoTrackService()
        {
            log4net.Config.XmlConfigurator.Configure();
            InitializeComponent();
            timer = new Timer();
        }

        public void OnDebug()
        {
            OnStart(null);
        }

        private void OnTimeElapsed(object source, ElapsedEventArgs e)
        {
            Logger.Info("Pdf Extract Process Started");
            AutoTrackPdf autoTrackPdf = new AutoTrackPdf();
            autoTrackPdf.Extract();
            Logger.Info("Pdf Extract Process Ended");


        }


        protected override void OnStart(string[] args)
        {
            Logger.Info("Pdf Extract Process Started");
            AutoTrackPdf autoTrackPdf = new AutoTrackPdf();
            autoTrackPdf.Extract();
            Logger.Info("Pdf Extract Process Ended");

            timer.Elapsed += new ElapsedEventHandler(OnTimeElapsed);
            timer.Interval = Convert.ToInt32(ConfigurationManager.AppSettings["ServiceRunAtEvery"]);
            timer.Enabled = true;
        }

        protected override void OnStop()
        {
            timer.Enabled = false;
            Logger.Info("Service Stoped.");
        }
    }
}

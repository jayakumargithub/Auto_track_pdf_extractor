using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Linq;
using System.ServiceProcess;
using System.Threading.Tasks;

namespace PdfDataExtractor
{
    [RunInstaller(true)]
    public partial class ProjectInstaller : System.Configuration.Install.Installer
    {
        public ProjectInstaller()
        {
            InitializeComponent();
            this.AfterInstall += ProjectInstaller_AfterInstall;
          //  this.AfterUninstall += ProjectInstaller_AfterUninstall;
        }

        private void ProjectInstaller_AfterUninstall(object sender, InstallEventArgs e)
        {
            using (ServiceController sc = new ServiceController(serviceInstaller1.ServiceName))
            { 
                sc.Stop();
            }
        }

 
          

        private void ProjectInstaller_AfterInstall(object sender, InstallEventArgs e)
        {
            using (ServiceController sc = new ServiceController(serviceInstaller1.ServiceName))
            {
                 
                    sc.Start(); 
            }
        }
    }
}

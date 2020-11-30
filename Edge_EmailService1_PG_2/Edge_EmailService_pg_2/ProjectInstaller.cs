using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Threading.Tasks;

namespace Edge_EmailService_pg
{
    [RunInstaller(true)]
    public partial class ProjectInstaller : System.Configuration.Install.Installer
    {
        protected static string AppPath = System.AppDomain.CurrentDomain.BaseDirectory;
        protected static string strLogFilePath = AppPath + "\\" + "Edge_EmailLog_pg.txt";
        private static StreamWriter sw = null;

        public ProjectInstaller()
        {
            InitializeComponent();
        }

        private void ProjectInstaller_AfterUninstall(object sender, InstallEventArgs e)
        {
            try
            {
                // Remove Event Source if already there    

                if (File.Exists(AppPath + "\\" + "Edge_EmailLog_pg.txt"))
                {
                    File.Delete(AppPath + "\\" + "Edge_EmailLog_pg.txt");
                }
            }
            catch (Exception ex)
            {
                LogExceptions(strLogFilePath, ex);
            }
        }

        private void ProjectInstaller_BeforeInstall(object sender, InstallEventArgs e)
        {
            try
            {
                ServiceController controller = new ServiceController(serviceInstallerEmailServicepg.ServiceName);
                if (controller.Status == ServiceControllerStatus.Running | controller.Status == ServiceControllerStatus.Paused )
                {
                    controller.Stop();
                    controller.WaitForStatus(ServiceControllerStatus.Stopped, new TimeSpan(0, 0, 0, 15));
                    controller.Close();
                }
            }
            catch (Exception ex)
            {
                LogExceptions(strLogFilePath, ex);
            }
        }

        public static void LogExceptions(string filePath, Exception ex)
        {
            if (false == File.Exists(filePath))
            {
                FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                fs.Close();
            }
            WriteExceptionLog(filePath, ex);
        }

        private static void WriteExceptionLog(string strPathName, Exception objException)
        {
            sw = new StreamWriter(strPathName, true);
            sw.WriteLine("Source		: " + objException.Source.ToString().Trim());
            sw.WriteLine("Method		: " + objException.TargetSite.Name.ToString());
            sw.WriteLine("Date		: " + DateTime.Now.ToLongTimeString());
            sw.WriteLine("Time		: " + DateTime.Now.ToShortDateString());
            sw.WriteLine("Error		: " + objException.Message.ToString().Trim());
            sw.WriteLine("Stack Trace	: " + objException.StackTrace.ToString().Trim());
            sw.WriteLine("^^-------------------------------------------------------------------^^");
            sw.Flush();
            sw.Close();
        }

        private void serviceInstallerEmailService_AfterInstall(object sender, InstallEventArgs e)
        {

        }
    }
}

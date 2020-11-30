namespace Edge_EmailService_pg
{
    partial class ProjectInstaller
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.serviceProcessInstallerEmailServicepg = new System.ServiceProcess.ServiceProcessInstaller();
            this.serviceInstallerEmailServicepg = new System.ServiceProcess.ServiceInstaller();
            // 
            // serviceProcessInstallerEmailServicepg
            // 
            this.serviceProcessInstallerEmailServicepg.Account = System.ServiceProcess.ServiceAccount.LocalService;
            this.serviceProcessInstallerEmailServicepg.Password = null;
            this.serviceProcessInstallerEmailServicepg.Username = null;
            // 
            // serviceInstallerEmailServicepg
            // 
            this.serviceInstallerEmailServicepg.DelayedAutoStart = true;
            this.serviceInstallerEmailServicepg.DisplayName = "OmegacubeEmailService_pg";
            this.serviceInstallerEmailServicepg.ServiceName = "OmegacubeEmailService_pg";
            this.serviceInstallerEmailServicepg.StartType = System.ServiceProcess.ServiceStartMode.Automatic;
            this.serviceInstallerEmailServicepg.AfterInstall += new System.Configuration.Install.InstallEventHandler(this.serviceInstallerEmailService_AfterInstall);
            // 
            // ProjectInstaller
            // 
            this.Installers.AddRange(new System.Configuration.Install.Installer[] {
            this.serviceProcessInstallerEmailServicepg,
            this.serviceInstallerEmailServicepg});
            this.AfterUninstall += new System.Configuration.Install.InstallEventHandler(this.ProjectInstaller_AfterUninstall);
            this.BeforeInstall += new System.Configuration.Install.InstallEventHandler(this.ProjectInstaller_BeforeInstall);

        }

        #endregion

        private System.ServiceProcess.ServiceProcessInstaller serviceProcessInstallerEmailServicepg;
        private System.ServiceProcess.ServiceInstaller serviceInstallerEmailServicepg;
    }
}
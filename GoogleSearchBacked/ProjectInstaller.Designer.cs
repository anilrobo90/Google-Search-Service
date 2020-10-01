namespace GoogleSearchBacked
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
            this.GoogleCustomSearchBackend = new System.ServiceProcess.ServiceProcessInstaller();
            this.serviceInstaller1 = new System.ServiceProcess.ServiceInstaller();
            // 
            // GoogleCustomSearchBackend
            // 
            this.GoogleCustomSearchBackend.Account = System.ServiceProcess.ServiceAccount.LocalSystem;
            this.GoogleCustomSearchBackend.Password = null;
            this.GoogleCustomSearchBackend.Username = null;
            // 
            // serviceInstaller1
            // 
            this.serviceInstaller1.Description = "Searchs and outputs Goole search results in excel file.";
            this.serviceInstaller1.DisplayName = "GoogleCustomSearchBackend";
            this.serviceInstaller1.ServiceName = "Service1";
            // 
            // ProjectInstaller
            // 
            this.Installers.AddRange(new System.Configuration.Install.Installer[] {
            this.GoogleCustomSearchBackend,
            this.serviceInstaller1});

        }

        #endregion

        private System.ServiceProcess.ServiceProcessInstaller GoogleCustomSearchBackend;
        private System.ServiceProcess.ServiceInstaller serviceInstaller1;
    }
}
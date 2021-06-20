namespace LWC.WallpaperChangerService
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
            this.WallpaperChangerProcessInstaller = new System.ServiceProcess.ServiceProcessInstaller();
            this.WallpaperChangerService = new System.ServiceProcess.ServiceInstaller();
            // 
            // WallpaperChangerProcessInstaller
            // 
            this.WallpaperChangerProcessInstaller.Account = System.ServiceProcess.ServiceAccount.LocalService;
            this.WallpaperChangerProcessInstaller.Password = null;
            this.WallpaperChangerProcessInstaller.Username = null;
            // 
            // WallpaperChangerService
            // 
            this.WallpaperChangerService.DelayedAutoStart = true;
            this.WallpaperChangerService.Description = "Manages the slideshow images available for desktop wallpapers.";
            this.WallpaperChangerService.DisplayName = "Wallpaper Changer Service";
            this.WallpaperChangerService.ServiceName = "Wallpaper Changer";
            this.WallpaperChangerService.StartType = System.ServiceProcess.ServiceStartMode.Automatic;
            // 
            // ProjectInstaller
            // 
            this.Installers.AddRange(new System.Configuration.Install.Installer[] {
            this.WallpaperChangerProcessInstaller,
            this.WallpaperChangerService});

        }

        #endregion

        private System.ServiceProcess.ServiceProcessInstaller WallpaperChangerProcessInstaller;
        private System.ServiceProcess.ServiceInstaller WallpaperChangerService;
    }
}
namespace GiaoTrinhWord.AddIn
{
    partial class RibbonGiaoTrinh : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonGiaoTrinh()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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
            Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher ribbonDialogLauncherImpl1 = this.Factory.CreateRibbonDialogLauncher();
            this.tabGiaoTrinh = this.Factory.CreateRibbonTab();
            this.groupGiaoTrinh = this.Factory.CreateRibbonGroup();
            this.btnMoGiaoTrinh = this.Factory.CreateRibbonButton();
            this.tabGiaoTrinh.SuspendLayout();
            this.groupGiaoTrinh.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabGiaoTrinh
            // 
            this.tabGiaoTrinh.Groups.Add(this.groupGiaoTrinh);
            this.tabGiaoTrinh.Label = "GIÁO TRÌNH";
            this.tabGiaoTrinh.Name = "tabGiaoTrinh";
            // 
            // groupGiaoTrinh
            // 
            this.groupGiaoTrinh.DialogLauncher = ribbonDialogLauncherImpl1;
            this.groupGiaoTrinh.Items.Add(this.btnMoGiaoTrinh);
            this.groupGiaoTrinh.Label = "Giáo Trình Word";
            this.groupGiaoTrinh.Name = "groupGiaoTrinh";
            // 
            // btnMoGiaoTrinh
            // 
            this.btnMoGiaoTrinh.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnMoGiaoTrinh.Image = global::GiaoTrinhWord.AddIn.Properties.Resources.icon_80;
            this.btnMoGiaoTrinh.Label = "Mở Giáo Trình";
            this.btnMoGiaoTrinh.Name = "btnMoGiaoTrinh";
            this.btnMoGiaoTrinh.ShowImage = true;
            this.btnMoGiaoTrinh.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnMoGiaoTrinh_Click_1);
            // 
            // RibbonGiaoTrinh
            // 
            this.Name = "RibbonGiaoTrinh";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tabGiaoTrinh);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonGiaoTrinh_Load);
            this.tabGiaoTrinh.ResumeLayout(false);
            this.tabGiaoTrinh.PerformLayout();
            this.groupGiaoTrinh.ResumeLayout(false);
            this.groupGiaoTrinh.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabGiaoTrinh;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupGiaoTrinh;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnMoGiaoTrinh;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonGiaoTrinh RibbonGiaoTrinh
        {
            get { return this.GetRibbon<RibbonGiaoTrinh>(); }
        }
    }
}

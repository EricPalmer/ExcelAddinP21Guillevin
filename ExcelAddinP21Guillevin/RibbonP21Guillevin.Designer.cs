
namespace ExcelAddinP21Guillevin {
    partial class RibbonP21Guillevin : Microsoft.Office.Tools.Ribbon.RibbonBase {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonP21Guillevin()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RibbonP21Guillevin));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.groupGuillevin = this.Factory.CreateRibbonGroup();
            this.groupSalesHistory = this.Factory.CreateRibbonGroup();
            this.groupMjolnir = this.Factory.CreateRibbonGroup();
            this.groupHelp = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.buttonFormatSalesHistory = this.Factory.CreateRibbonButton();
            this.btnMjolnirNewPage = this.Factory.CreateRibbonButton();
            this.btnMjolnirRun = this.Factory.CreateRibbonButton();
            this.btnHelp = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.groupGuillevin.SuspendLayout();
            this.groupSalesHistory.SuspendLayout();
            this.groupMjolnir.SuspendLayout();
            this.groupHelp.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.groupGuillevin);
            this.tab1.Groups.Add(this.groupSalesHistory);
            this.tab1.Groups.Add(this.groupMjolnir);
            this.tab1.Groups.Add(this.groupHelp);
            this.tab1.Label = "Guillevin - P21";
            this.tab1.Name = "tab1";
            // 
            // groupGuillevin
            // 
            this.groupGuillevin.Items.Add(this.button1);
            this.groupGuillevin.Label = "ver 0.4";
            this.groupGuillevin.Name = "groupGuillevin";
            // 
            // groupSalesHistory
            // 
            this.groupSalesHistory.Items.Add(this.buttonFormatSalesHistory);
            this.groupSalesHistory.Label = "Sales History";
            this.groupSalesHistory.Name = "groupSalesHistory";
            // 
            // groupMjolnir
            // 
            this.groupMjolnir.Items.Add(this.btnMjolnirNewPage);
            this.groupMjolnir.Items.Add(this.btnMjolnirRun);
            this.groupMjolnir.Label = "Mjölnir";
            this.groupMjolnir.Name = "groupMjolnir";
            // 
            // groupHelp
            // 
            this.groupHelp.Items.Add(this.btnHelp);
            this.groupHelp.Name = "groupHelp";
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
            this.button1.Label = " ";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            // 
            // buttonFormatSalesHistory
            // 
            this.buttonFormatSalesHistory.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonFormatSalesHistory.Label = "Format";
            this.buttonFormatSalesHistory.Name = "buttonFormatSalesHistory";
            this.buttonFormatSalesHistory.OfficeImageId = "MorePagePartsInsert";
            this.buttonFormatSalesHistory.ShowImage = true;
            this.buttonFormatSalesHistory.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonFormatSalesHistory_Click);
            // 
            // btnMjolnirNewPage
            // 
            this.btnMjolnirNewPage.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnMjolnirNewPage.Label = "New Page";
            this.btnMjolnirNewPage.Name = "btnMjolnirNewPage";
            this.btnMjolnirNewPage.OfficeImageId = "SlideNew";
            this.btnMjolnirNewPage.ShowImage = true;
            this.btnMjolnirNewPage.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnMjolnirNewPage_Click);
            // 
            // btnMjolnirRun
            // 
            this.btnMjolnirRun.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnMjolnirRun.Image = ((System.Drawing.Image)(resources.GetObject("btnMjolnirRun.Image")));
            this.btnMjolnirRun.Label = "Run";
            this.btnMjolnirRun.Name = "btnMjolnirRun";
            this.btnMjolnirRun.ShowImage = true;
            this.btnMjolnirRun.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnMjolnirRun_Click);
            // 
            // btnHelp
            // 
            this.btnHelp.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnHelp.Label = "Help";
            this.btnHelp.Name = "btnHelp";
            this.btnHelp.OfficeImageId = "Help";
            this.btnHelp.ShowImage = true;
            this.btnHelp.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnHelp_Click);
            // 
            // RibbonP21Guillevin
            // 
            this.Name = "RibbonP21Guillevin";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonP21Guillevin_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.groupGuillevin.ResumeLayout(false);
            this.groupGuillevin.PerformLayout();
            this.groupSalesHistory.ResumeLayout(false);
            this.groupSalesHistory.PerformLayout();
            this.groupMjolnir.ResumeLayout(false);
            this.groupMjolnir.PerformLayout();
            this.groupHelp.ResumeLayout(false);
            this.groupHelp.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupSalesHistory;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonFormatSalesHistory;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupGuillevin;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupMjolnir;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnMjolnirNewPage;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupHelp;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnHelp;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnMjolnirRun;
    }

    partial class ThisRibbonCollection {
        internal RibbonP21Guillevin RibbonP21Guillevin
        {
            get { return this.GetRibbon<RibbonP21Guillevin>(); }
        }
    }
}

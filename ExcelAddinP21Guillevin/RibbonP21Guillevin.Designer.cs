
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
            this.group2 = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.buttonFormatSalesHistory = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.btnMjolnirNewPage = this.Factory.CreateRibbonButton();
            this.btnMjolnirRun = this.Factory.CreateRibbonButton();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.btnHelp = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group1.SuspendLayout();
            this.group3.SuspendLayout();
            this.group4.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.group4);
            this.tab1.Label = "Guillevin - P21";
            this.tab1.Name = "tab1";
            // 
            // group2
            // 
            this.group2.Items.Add(this.button1);
            this.group2.Label = "ver 0.2";
            this.group2.Name = "group2";
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
            this.button1.Label = " ";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            // 
            // group1
            // 
            this.group1.Items.Add(this.buttonFormatSalesHistory);
            this.group1.Label = "Sales History";
            this.group1.Name = "group1";
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
            // group3
            // 
            this.group3.Items.Add(this.btnMjolnirNewPage);
            this.group3.Items.Add(this.btnMjolnirRun);
            this.group3.Label = "Mjölnir";
            this.group3.Name = "group3";
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
            // group4
            // 
            this.group4.Items.Add(this.btnHelp);
            this.group4.Name = "group4";
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
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonFormatSalesHistory;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnMjolnirNewPage;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
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

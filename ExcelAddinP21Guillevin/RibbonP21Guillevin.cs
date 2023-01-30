using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Windows.Forms;
using System.ComponentModel;



namespace ExcelAddinP21Guillevin {
    public partial class RibbonP21Guillevin {
        //BackgroundWorker WorkerThread;
        BackgroundWorker backgroundWorker;
        public LoadingForm frmLoadingForm;
        FormatSalesHistory formatSalesHistory;

        private void RibbonP21Guillevin_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void buttonFormatSalesHistory_Click(object sender, RibbonControlEventArgs e)
        {
            // Create FormatSalesHistory object to use in background task
            formatSalesHistory = new FormatSalesHistory();

            frmLoadingForm = new LoadingForm();

            System.Threading.SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
            backgroundWorker = new BackgroundWorker();
            backgroundWorker.WorkerReportsProgress = true;
            backgroundWorker.DoWork += new DoWorkEventHandler(FormatSalesHistory_BGWorker_DoWork);
            backgroundWorker.ProgressChanged += new ProgressChangedEventHandler(FormatSalesHistory_BGWorker_ProgressChanged);
            backgroundWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(FormatSalesHistory_BGWorker_RunWorkerCompleted);
            backgroundWorker.RunWorkerAsync();

            frmLoadingForm.ShowDialog();
        }

        private void FormatSalesHistory_BGWorker_DoWork(object sender, DoWorkEventArgs e) {
            // Parse the sales history data on the active worksheet
            formatSalesHistory.Format(Globals.ThisAddIn.Application.ActiveSheet, backgroundWorker);
        }

        private void FormatSalesHistory_BGWorker_ProgressChanged(object sender, ProgressChangedEventArgs e) {
            // Update progress bar
            frmLoadingForm.updateProgressBar(e.ProgressPercentage);
        }

        private void FormatSalesHistory_BGWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e) {
            // Create new worksheet to put parsed data
            formatSalesHistory.FormatWorksheet(Globals.ThisAddIn.Application.ActiveSheet);

            // Close loading form
            frmLoadingForm.Close();
        }





        private void btnMjolnirNewPage_Click(object sender, RibbonControlEventArgs e) {
            Mjolnir mjolnir = new Mjolnir();
            mjolnir.NewPage(Globals.ThisAddIn.Application.ActiveSheet);
        }

        private void btnMjolnirRun_Click(object sender, RibbonControlEventArgs e) {
            Mjolnir mjolnir = new Mjolnir();
            mjolnir.Run(Globals.ThisAddIn.Application.ActiveSheet);
        }
        
        private void btnHelp_Click(object sender, RibbonControlEventArgs e) {
            Globals.ThisAddIn.taskPane.Visible = !Globals.ThisAddIn.taskPane.Visible;
        }
    }
}

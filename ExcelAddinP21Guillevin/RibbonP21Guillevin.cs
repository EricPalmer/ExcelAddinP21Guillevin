using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Windows.Forms;
using System.ComponentModel;
using System.Reflection;



namespace ExcelAddinP21Guillevin {
    public partial class RibbonP21Guillevin {
        //BackgroundWorker WorkerThread;
        BackgroundWorker backgroundWorker;
        public LoadingForm frmLoadingForm;
        FormatSalesHistory formatSalesHistory;

        private void RibbonP21Guillevin_Load(object sender, RibbonUIEventArgs e)
        {
            Version ver;
            if (System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed) {
                ver = System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion;
                groupGuillevin.Label = string.Format("Ver {0}.{1}.{2}.{3}", ver.Major, ver.Minor, ver.Build, ver.Revision);
            }
            else {
                ver = Assembly.GetExecutingAssembly().GetName().Version;
                groupGuillevin.Label = string.Format("Ver {0}.{1}.{2}.{3}", ver.Major, ver.Minor, ver.Build, ver.Revision);
            }
        }

        private void buttonFormatSalesHistory_Click(object sender, RibbonControlEventArgs e)
        {
            // Check if this is a valid sales history page
            if (Globals.ThisAddIn.Application.ActiveSheet.Cells[1,2].Value2 != "Detailed Sales History Report") {
                MessageBox.Show("This is not a valid sales history report sheet.");
                return;
            }

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
            // Create FormatSalesHistory object to use in background task
            // Need to do this because excel doesn't like it when you access it from a background worker
            // So the constructor copies all the cells into an array for its internal functions to use instead

            formatSalesHistory = new FormatSalesHistory(Globals.ThisAddIn.Application.ActiveSheet);

            // Parse the sales history data on the active worksheet
            formatSalesHistory.Parse(backgroundWorker);
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
            //Globals.ThisAddIn.taskPane.Visible = !Globals.ThisAddIn.taskPane.Visible;
        }
    }
}

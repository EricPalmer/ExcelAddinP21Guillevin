using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;



namespace ExcelAddinP21Guillevin {
    public partial class RibbonP21Guillevin {
        private void RibbonP21Guillevin_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void buttonFormatSalesHistory_Click(object sender, RibbonControlEventArgs e)
        {
            FormatSalesHistory formatSalesHistory = new FormatSalesHistory(Globals.ThisAddIn.Application.ActiveSheet);
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

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;

namespace ExcelAddinP21Guillevin
{
    public partial class ThisAddIn
    {
        private Microsoft.Office.Tools.CustomTaskPane actionPane;
        private ActionPaneHelp actionPaneHelp;

        public Microsoft.Office.Tools.CustomTaskPane taskPane {
            get {
                return actionPane;
            }
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e) {
            // Create a new instance of the user control
            //actionPaneHelp = new ActionPaneHelp();

            // Create a new CustomTaskPane object and add the user control to it
            //actionPane = this.CustomTaskPanes.Add(actionPaneHelp, "My Action Pane");

            //actionPane.Visible = true;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e) {


        }


        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}

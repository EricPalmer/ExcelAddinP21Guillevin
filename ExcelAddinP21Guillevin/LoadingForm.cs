using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelAddinP21Guillevin {
    public partial class LoadingForm : Form {
        public LoadingForm() {
            InitializeComponent();
        }

        public void updateProgressBar(int percentage) {
            LoadingBar.Value = percentage;
        }

    }
}

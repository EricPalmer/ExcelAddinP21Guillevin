using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace ExcelAddinP21Guillevin {
    class Mjolnir {
        struct product {
            public string productName;
            public string quantity;
        }
        public Mjolnir() { }

        public void NewPage(Excel.Worksheet wsCur) {
            // Create a new worksheet and set up the headers
            Excel.Worksheet wsNew = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(After: wsCur);

            wsNew.Cells[1, 1] = "Product Name";
            wsNew.Cells[1, 2] = "Quantity";

            // Bold the headers
            wsNew.Range["A1:B1"].Font.Bold = true;

            // Freeze the top row
            wsNew.Application.ActiveWindow.SplitRow = 1;
            wsNew.Application.ActiveWindow.FreezePanes = true;

            // Resize columns
            wsNew.UsedRange.Columns.AutoFit();
        }

        public void Run(Excel.Worksheet wsCur) {
            try {
                // First, check if we are on a mjolnir page
                // If not, exit
                if (!(Convert.ToString(wsCur.Cells[1, 1].Value2) == "Product Name")) {
                    MessageBox.Show("Please switch to the product sheet before running this");
                    return;
                }

                // Get list of products from sheet
                List<product> products = new List<product>();
                ParseData(wsCur, products);

                // Paste products one by one in P21
                WriteData(products);
            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message);
                return;
            }
        }            

        private void ParseData(Excel.Worksheet ws, List<product> products) { 
            // Get total row count
            long rowCount = ws.UsedRange.Rows.Count + ws.UsedRange.Rows[1].Row - 1;

            product product = new product();

            for (long curRow = 2; curRow <= rowCount; curRow++) {
                // Grab the values in the current row
                // Validate them
                product.productName = Convert.ToString(ws.Cells[curRow, 1].Value2);
                if (product.productName == null) {
                    product.productName = "";
                }

                product.quantity = Convert.ToString(ws.Cells[curRow, 2].Value2);
                if (product.quantity == null) {
                    product.quantity = "";
                }

                // If both are blank, do nothing and continue
                if (product.productName == "" && product.quantity == "") {
                    continue;
                }

                // If the quantity value is not a number, throw an error and exit
                int quantity;
                if (!int.TryParse(product.quantity, out quantity)) {
                    throw new InvalidOperationException("Quantity is invalid on row " + curRow + ".");
                }

                // If only one of the columns has data in it, but the other doesn't, throw an error and exit
                if (product.productName == "" ^ product.quantity == "") {
                    throw new InvalidOperationException("Missing information on row " + curRow + ".");
                }

                products.Add(product);
            }
        }

        private void WriteData(List<product> products) {
            // Let the user know that they will have 5 seconds to make it to P21
            DialogResult result = MessageBox.Show("Once you click OK you will have 7 seconds to navigate to the lightning bolt", "Mjölnir", MessageBoxButtons.OKCancel);
            if (result == DialogResult.OK) {
                System.Threading.Thread.Sleep(7000);

                // Paste the product name and quantity, hitting enter after each one to fill in the P21 order entry
                foreach (product product in products) {
                    SendKeys.Send(product.productName);
                    System.Threading.Thread.Sleep(500);
                    SendKeys.Send("{ENTER}");
                    System.Threading.Thread.Sleep(500);

                    SendKeys.Send(product.quantity);
                    System.Threading.Thread.Sleep(500);
                    SendKeys.Send("{ENTER}");
                    System.Threading.Thread.Sleep(1000);
                }
            }
            else {
                return;
            }
        }
    }
}

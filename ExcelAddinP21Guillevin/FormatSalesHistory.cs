using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.ComponentModel;

namespace ExcelAddinP21Guillevin {
    public class FormatSalesHistory {
        struct SalesHistoryEntry {
            public string branchNum;
            public string branchName;
            public string salesRep;
            public string invoiceDate;
            public string customerName;
            public string customerPostalCode;
            public string partNumber;
            public string partDesc;
            public double quantity;
            public double totalCost;
            public double unitCost;
            public double totalPrice;
            public double unitPrice;
        }

        List<SalesHistoryEntry> entries;

        public FormatSalesHistory() { 
        
        }
        public void Parse(Excel.Worksheet DataWs, BackgroundWorker backgroundWorker)
        {
            try {
                // Get all entries from given worksheet
                entries = new List<SalesHistoryEntry>();

                // Get total row count
                long rowCount = DataWs.UsedRange.Rows.Count + DataWs.UsedRange.Rows[1].Row - 1;

                // Copy used range to an array to process faster
                object[,] cellsArray = DataWs.get_Range("A1:Z" + rowCount).Value2;

                ParseData(cellsArray, rowCount, entries, backgroundWorker);
            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message);
                return;
            }
        }

        private void ParseData(object[,] cellsArray, long rowCount, List<SalesHistoryEntry> entries, BackgroundWorker backgroundWorker) {
            SalesHistoryEntry entry = new SalesHistoryEntry();

            //Loop from top to bottom
            //Search for an "E" or "C" in the P column, this is where an item detail is
            //Then go back up the excel file to fill in the rest of the details
            string curRowText;
            for (long curRow = 1; curRow < rowCount; curRow++) {
                // Update progress bar
                int progressPercentage = (int)((curRow + 1) * 100 / rowCount);
                backgroundWorker.ReportProgress(progressPercentage);

                curRowText = Convert.ToString(cellsArray[curRow, 16]);
                if (curRowText == null) {
                    curRowText = "";
                }

                if (curRowText == "E" || curRowText == "C") {
                    // If the row above contains a part number, grab it
                    // Otherwise, use the previous part number found
                    curRowText = Convert.ToString(cellsArray[curRow -1, 1]);
                    if (curRowText == null) {
                        curRowText = "";
                    }
                    if (!(curRowText == "")) {
                        entry.partNumber = Convert.ToString(cellsArray[curRow - 1, 1]).Trim();
                        entry.partDesc = Convert.ToString(cellsArray[curRow - 1, 3]).Trim();
                    }

                    entry.quantity = (double)cellsArray[curRow, 15];

                    entry.totalCost = (double)cellsArray[curRow, 18];
                    entry.unitCost = Math.Abs(entry.totalCost) / entry.quantity;

                    entry.totalPrice = (double)cellsArray[curRow, 17];
                    entry.unitPrice = Math.Abs(entry.totalPrice) / entry.quantity;

                    GetCustomerInfo(cellsArray, curRow, ref entry.customerName, ref entry.customerPostalCode);
                    GetInvoiceDate(cellsArray, curRow, ref entry.invoiceDate);
                    GetSalesLocation(cellsArray, curRow, ref entry.branchName, ref entry.branchNum);
                    GetSalesRep(cellsArray, curRow, ref entry.salesRep);

                    entries.Add(entry);
                }
            }
        }

        private void GetCustomerInfo(object[,] cellsArray, long curRow, ref string customerName, ref string customerPostalCode) {
            string inputText;
            int startIndex;
            int endIndex;

            // Loop back up until a cell is found that starts with "Customer : "
            do {
                curRow--;
                
                if (curRow == 0) {
                    throw new InvalidOperationException("Invalid format. Rolled back to the start of the excel file.");
                }

                inputText = Convert.ToString(cellsArray[curRow, 1]);
                if (inputText == null) {
                    inputText = "";
                }
            }
            while (!(inputText.Substring(0, inputText.Length > 11 ? 11 : inputText.Length) == "Customer : "));

            // Extract customer name
            startIndex = inputText.IndexOf('-') + 1;
            endIndex = inputText.IndexOf(',');
            customerName = inputText.Substring(startIndex, endIndex - startIndex).Trim();

            // Extract customer postal code
            startIndex = inputText.LastIndexOf(',') + 1;
            customerPostalCode = inputText.Substring(startIndex).Trim();
        }

        private void GetInvoiceDate(object[,] cellsArray, long curRow, ref string invoiceDate) {
            string inputText;
            int startIndex;

            // Loop back up until a cell is found that starts with "Invoice Date : "
            do {
                curRow--;

                if (curRow == 0) {
                    throw new InvalidOperationException("Invalid format. Rolled back to the start of the excel file.");
                }

                inputText = Convert.ToString(cellsArray[curRow, 1]);
                if (inputText == null) {
                    inputText = "";
                }
            }
            while (!(inputText.Substring(0, inputText.Length > 15 ? 15 : inputText.Length) == "Invoice Date : "));

            // Extract the invoice date
            startIndex = inputText.LastIndexOf(':') + 1;
            invoiceDate = inputText.Substring(startIndex).Trim();
        }

        private void GetSalesLocation(object[,] cellsArray, long curRow, ref string branchName, ref string branchNum) {
            string inputText;
            int startIndex;
            int endIndex;

            // Loop back up until a cell is found that starts with "Sales Location : "
            do {
                curRow--;

                if (curRow == 0) {
                    throw new InvalidOperationException("Invalid format. Rolled back to the start of the excel file.");
                }

                inputText = Convert.ToString(cellsArray[curRow, 1]);
                if (inputText == null) {
                    inputText = "";
                }
            }
            while (!(inputText.Substring(0, inputText.Length > 17 ? 17 : inputText.Length) == "Sales Location : "));

            // Extract branch name
            // First just grab everything after the '-'
            startIndex = inputText.IndexOf('-');
            branchName = inputText.Substring(startIndex);

            // Then check if there are '( )' in the line, if there are update the name with what is between them
            startIndex = inputText.IndexOf('(');
            endIndex = inputText.LastIndexOf(')');

            if (startIndex != -1 && endIndex != -1 && startIndex < endIndex) {
                startIndex++;
                branchName = inputText.Substring(startIndex, endIndex - startIndex).Trim();
            }

            // Extract branch num
            startIndex = inputText.IndexOf(':') + 1;
            endIndex = inputText.IndexOf('-');
            branchNum = inputText.Substring(startIndex, endIndex - startIndex).Trim();
        }

        private void GetSalesRep(object[,] cellsArray, long curRow, ref string salesRep) {
            string inputText;
            int startIndex;

            // Loop back up until a cell is found that starts with "SalesRep : "
            do {
                curRow--;

                if (curRow == 0) {
                    throw new InvalidOperationException("Invalid format. Rolled back to the start of the excel file.");
                }

                inputText = Convert.ToString(cellsArray[curRow, 1]);
                if (inputText == null) {
                    inputText = "";
                }
            }
            while (!(inputText.Substring(0, inputText.Length > 11 ? 11 : inputText.Length) == "SalesRep : "));

            // Extract the invoice date
            startIndex = inputText.LastIndexOf('-') + 1;
            salesRep = inputText.Substring(startIndex).Trim();
        }

        public void FormatWorksheet(Excel.Worksheet DataWs) {
            // Create new worksheet for data to be placed on
            Excel.Worksheet formattedWs = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(After: DataWs);

            // Put data into an array that will be copied in one go into the excel worksheet
            object[,] OutputArray = new object[entries.Count + 1, 13];

            OutputArray[0, 0] = "Branch Num";
            OutputArray[0, 1] = "Branch Name";
            OutputArray[0, 2] = "Sales Rep";
            OutputArray[0, 3] = "Invoice Date";
            OutputArray[0, 4] = "Customer";
            OutputArray[0, 5] = "Customer Postal Code";
            OutputArray[0, 6] = "Part Number";
            OutputArray[0, 7] = "Part Description";
            OutputArray[0, 8] = "Qty";
            OutputArray[0, 9] = "Total Cost";
            OutputArray[0, 10] = "Unit Cost";
            OutputArray[0, 11] = "Total Price";
            OutputArray[0, 12] = "Unit Price";

            // Add each entry to a new line on the sheet
            long curRow = 1;
            foreach (SalesHistoryEntry entry in entries) {
                OutputArray[curRow, 0] = entry.branchNum;
                OutputArray[curRow, 1] = entry.branchName;
                OutputArray[curRow, 2] = entry.salesRep;
                OutputArray[curRow, 3] = entry.invoiceDate;
                OutputArray[curRow, 4] = entry.customerName;
                OutputArray[curRow, 5] = entry.customerPostalCode;
                OutputArray[curRow, 6] = entry.partNumber;
                OutputArray[curRow, 7] = entry.partDesc;
                OutputArray[curRow, 8] = entry.quantity;
                OutputArray[curRow, 9] = entry.totalCost;
                OutputArray[curRow, 10] = entry.unitCost;
                OutputArray[curRow, 11] = entry.totalPrice;
                OutputArray[curRow, 12] = entry.unitPrice;

                curRow++;
            }

            // Copy the array contents onto the newly created worksheet
            formattedWs.Select();
            formattedWs.Range["A1", formattedWs.Cells[entries.Count, 13]] = OutputArray;

            // Bold the headers
            formattedWs.Range["A1:M1"].Font.Bold = true;

            // Freeze the top row
            formattedWs.Application.ActiveWindow.SplitRow = 1;
            formattedWs.Application.ActiveWindow.FreezePanes = true;

            // Format the currency cells
            formattedWs.Range["J2", formattedWs.Cells[entries.Count, 13]].Numberformat = "$0.00";

            // Resize columns
            formattedWs.UsedRange.Columns.AutoFit();
        }
    }
}

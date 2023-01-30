using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

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
        public FormatSalesHistory(Excel.Worksheet DataWs)
        {
            try {
                // Get all entries from given worksheet
                List<SalesHistoryEntry> entries = new List<SalesHistoryEntry>();
                ParseData(DataWs, entries);

                // Create new worksheet 
                Excel.Worksheet formattedWs = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(After: DataWs);
                FormatWorksheet(formattedWs, entries);
            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message);
                return;
            }
        }

        private void ParseData(Excel.Worksheet ws, List<SalesHistoryEntry> entries) {
            SalesHistoryEntry entry = new SalesHistoryEntry();

            // Get total row count
            long rowCount = ws.UsedRange.Rows.Count + ws.UsedRange.Rows[1].Row - 1;
            string curRowText;

            //Loop from top to bottom
            //Search for an "E" or "C" in the P column, this is where an item detail is
            //Then go back up the excel file to fill in the rest of the details
            for (long curRow = 1; curRow < rowCount; curRow++) {
                curRowText = Convert.ToString(ws.Cells[curRow, 16].Value2);
                if (curRowText == null) {
                    curRowText = "";
                }


                if (curRowText == "E" || curRowText == "C") {
                    // If the row above contains a part number, grab it
                    // Otherwise, use the previous part number found
                    curRowText = Convert.ToString(ws.Cells[curRow -1, 1].Value2);
                    if (curRowText == null) {
                        curRowText = "";
                    }
                    if (!(curRowText == "")) {
                        entry.partNumber = Convert.ToString(ws.Cells[curRow - 1, 1].Value2).Trim();
                        entry.partDesc = Convert.ToString(ws.Cells[curRow - 1, 3].Value2).Trim();
                    }

                    entry.quantity = ws.Cells[curRow, 15].Value2;

                    entry.totalCost = ws.Cells[curRow, 18].Value2;
                    entry.unitCost = Math.Abs(entry.totalCost) / entry.quantity;

                    entry.totalPrice = ws.Cells[curRow, 17].Value2;
                    entry.unitPrice = Math.Abs(entry.totalPrice) / entry.quantity;

                    GetCustomerInfo(ws, curRow, ref entry.customerName, ref entry.customerPostalCode);
                    GetInvoiceDate(ws, curRow, ref entry.invoiceDate);
                    GetSalesLocation(ws, curRow, ref entry.branchName, ref entry.branchNum);
                    GetSalesRep(ws, curRow, ref entry.salesRep);

                    entries.Add(entry);
                }
            }
        }

        private void GetCustomerInfo(Excel.Worksheet ws, long curRow, ref string customerName, ref string customerPostalCode) {
            string inputText;
            int startIndex;
            int endIndex;

            // Loop back up until a cell is found that starts with "Customer : "
            do {
                curRow--;
                
                if (curRow == 0) {
                    throw new InvalidOperationException("Invalid format. Rolled back to the start of the excel file.");
                }

                inputText = Convert.ToString(ws.Cells[curRow, 1].Value2);
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

        private void GetInvoiceDate(Excel.Worksheet ws, long curRow, ref string invoiceDate) {
            string inputText;
            int startIndex;

            // Loop back up until a cell is found that starts with "Invoice Date : "
            do {
                curRow--;

                if (curRow == 0) {
                    throw new InvalidOperationException("Invalid format. Rolled back to the start of the excel file.");
                }

                inputText = Convert.ToString(ws.Cells[curRow, 1].Value2);
                if (inputText == null) {
                    inputText = "";
                }
            }
            while (!(inputText.Substring(0, inputText.Length > 15 ? 15 : inputText.Length) == "Invoice Date : "));

            // Extract the invoice date
            startIndex = inputText.LastIndexOf(':') + 1;
            invoiceDate = inputText.Substring(startIndex).Trim();
        }

        private void GetSalesLocation(Excel.Worksheet ws, long curRow, ref string branchName, ref string branchNum) {
            string inputText;
            int startIndex;
            int endIndex;

            // Loop back up until a cell is found that starts with "Sales Location : "

            do {
                curRow--;

                if (curRow == 0) {
                    throw new InvalidOperationException("Invalid format. Rolled back to the start of the excel file.");
                }

                inputText = Convert.ToString(ws.Cells[curRow, 1].Value2);
                if (inputText == null) {
                    inputText = "";
                }
            }
            while (!(inputText.Substring(0, inputText.Length > 17 ? 17 : inputText.Length) == "Sales Location : "));

            // Extract branch name
            startIndex = inputText.IndexOf('(') + 1;
            endIndex = inputText.LastIndexOf(')');
            branchName = inputText.Substring(startIndex, endIndex - startIndex).Trim();

            // Extract branch num
            startIndex = inputText.IndexOf(':') + 1;
            endIndex = inputText.IndexOf('-');
            branchNum = inputText.Substring(startIndex, endIndex - startIndex).Trim();
        }

        private void GetSalesRep(Excel.Worksheet ws, long curRow, ref string salesRep) {
            string inputText;
            int startIndex;

            // Loop back up until a cell is found that starts with "SalesRep : "
            do {
                curRow--;

                if (curRow == 0) {
                    throw new InvalidOperationException("Invalid format. Rolled back to the start of the excel file.");
                }

                inputText = Convert.ToString(ws.Cells[curRow, 1].Value2);
                if (inputText == null) {
                    inputText = "";
                }
            }
            while (!(inputText.Substring(0, inputText.Length > 11 ? 11 : inputText.Length) == "SalesRep : "));

            // Extract the invoice date
            startIndex = inputText.LastIndexOf('-') + 1;
            salesRep = inputText.Substring(startIndex).Trim();
        }

        private void FormatWorksheet(Excel.Worksheet ws, List<SalesHistoryEntry> entries) {
            ws.Select();

            ws.Cells[1, 1] = "Branch Num";
            ws.Cells[1, 2] = "Branch Name";
            ws.Cells[1, 3] = "Sales Rep";
            ws.Cells[1, 4] = "Invoice Date";
            ws.Cells[1, 5] = "Customer";
            ws.Cells[1, 6] = "Customer Postal Code";
            ws.Cells[1, 7] = "Part Number";
            ws.Cells[1, 8] = "Part Description";
            ws.Cells[1, 9] = "Qty";
            ws.Cells[1, 10] = "Total Cost";
            ws.Cells[1, 11] = "Unit Cost";
            ws.Cells[1, 12] = "Total Price";
            ws.Cells[1, 13] = "Unit Price";

            // Bold the headers
            ws.Range["A1:M1"].Font.Bold = true;

            // Freeze the top row
            ws.Application.ActiveWindow.SplitRow = 1;
            ws.Application.ActiveWindow.FreezePanes = true;

            // Add each entry to a new line on the sheet
            long curRow = 2;
            foreach (SalesHistoryEntry entry in entries) {
                ws.Cells[curRow, 1] = entry.branchNum;
                ws.Cells[curRow, 2] = entry.branchName;
                ws.Cells[curRow, 3] = entry.salesRep;
                ws.Cells[curRow, 4] = entry.invoiceDate;
                ws.Cells[curRow, 5] = entry.customerName;
                ws.Cells[curRow, 6] = entry.customerPostalCode;
                ws.Cells[curRow, 7] = entry.partNumber;
                ws.Cells[curRow, 8] = entry.partDesc;
                ws.Cells[curRow, 9] = entry.quantity;
                ws.Cells[curRow, 10] = entry.totalCost;
                ws.Cells[curRow, 10].NumberFormat = "$#,##0.00";
                ws.Cells[curRow, 11] = entry.unitCost;
                ws.Cells[curRow, 11].Numberformat = "$0.00";
                ws.Cells[curRow, 12] = entry.totalPrice;
                ws.Cells[curRow, 12].Numberformat = "$0.00";
                ws.Cells[curRow, 13] = entry.unitPrice;
                ws.Cells[curRow, 13].Numberformat = "$0.00";

                curRow++;
            }

            // Resize columns
            ws.UsedRange.Columns.AutoFit();
        }
    }
}

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
        // Dictionary to hold a list of all the sort levels and their associated function to extract that particular piece of data
        Dictionary<string, Action<object[,], long, SalesHistoryEntry>> sortLevelReadFunctions;

        // List of the sort levels used in the generated report
        List<String> sortLevels;

        class SalesHistoryEntry {
            public string branchNum;
            public string branchName;
            public string salesRep;
            public string invoicePeriod;
            public string invoiceDate;
            public string invoiceNum;
            public string orderNum;
            public string customerName;
            public string customerPostalCode;
            public string customerShipToNum;
            public string partNumber;
            public string partDesc;
            public double quantity;
            public double totalCost;
            public double unitCost;
            public double totalPrice;
            public double unitPrice;
            public double marginPercent;
            public string contractNum;
            public string productGroup;
            public string supplier;
            public string taker;
        }

        List<SalesHistoryEntry> entries;
        object[,] cellsArray;
        long rowCount;

        public FormatSalesHistory(Excel.Worksheet DataWs) {
            // Get all entries from given worksheet
            entries = new List<SalesHistoryEntry>();

            // Get total row count
            rowCount = DataWs.UsedRange.Rows.Count + DataWs.UsedRange.Rows[1].Row - 1;

            // Copy used range to an array to process faster
            cellsArray = DataWs.get_Range("A1:AZ" + rowCount).Value2;
        }

        public void Parse(BackgroundWorker backgroundWorker) {
            try {
                sortLevelReadFunctions = new Dictionary<string, Action<object[,], long, SalesHistoryEntry>> {
                    {"Customer",  new Action<object[,], long, SalesHistoryEntry>(GetCustomerInfo)},
                    {"Item",  new Action<object[,], long, SalesHistoryEntry>(GetItemInfo)},
                    {"Invoice Date",  new Action<object[,], long, SalesHistoryEntry>(GetInvoiceDateInfo)},
                    {"Sales Location",  new Action<object[,], long, SalesHistoryEntry>(GetSalesLocationInfo)},
                    {"Sales Rep",  new Action<object[,], long, SalesHistoryEntry>(GetSalesRepInfo)},
                    {"Branch",  new Action<object[,], long, SalesHistoryEntry>(GetBranchInfo)},
                    {"Contract Number",  new Action<object[,], long, SalesHistoryEntry>(GetContractNumberInfo)},
                    {"Invoice Period",  new Action<object[,], long, SalesHistoryEntry>(GetInvoicePeriodInfo)},
                    {"Product Group",  new Action<object[,], long, SalesHistoryEntry>(GetProductGroupInfo)},
                    {"Ship To",  new Action<object[,], long, SalesHistoryEntry>(GetShipToInfo)},
                    {"Supplier",  new Action<object[,], long, SalesHistoryEntry>(GetSupplierInfo)},
                    {"Taker",  new Action<object[,], long, SalesHistoryEntry>(GetTakerInfo)},
                };

                // Get list of up to 5 'Sort Levels'
                sortLevels = new List<string>();
                int sortLevelCol = 8;

                // Get the correct column for where the current sort level is located: 8, 14, 19, 25, 31
                for (int curSortLevel = 0; curSortLevel < 5; curSortLevel++) {
                    switch (curSortLevel) {
                        case 0:
                            sortLevelCol = 8;
                            break;
                        case 1:
                            sortLevelCol = 13;
                            break;
                        case 2:
                            sortLevelCol = 19;
                            break;
                        case 3:
                            sortLevelCol = 25;
                            break;
                        case 4:
                            sortLevelCol = 31;
                            break;
                    }

                    string sortLevel = Convert.ToString(cellsArray[1, sortLevelCol]);
                    if (sortLevel == null) {
                        sortLevel = "";
                    }

                    if (sortLevel != "None") {
                        // Check if this sortLevel is in the dictionary
                        // If not, throw an error
                        if (!sortLevelReadFunctions.ContainsKey(sortLevel)) {
                            throw new InvalidOperationException("New sort level found. Please send this file to Eric to fix!");
                        }

                        sortLevels.Add(sortLevel);
                    }
                }

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
                        SalesHistoryEntry entry = new SalesHistoryEntry();

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
                        else {
                            entry.partNumber = entries.Last().partNumber;
                            entry.partDesc = entries.Last().partDesc;
                        }

                        // Loop through each of the sort levels and call the associated getInfo function
                        foreach (string sortlevel in sortLevels) {
                            sortLevelReadFunctions[sortlevel](cellsArray, curRow, entry);
                        }

                        entries.Add(entry);
                    }
                }
            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message);
                return;
            }
        }

        private static void GetCustomerInfo(object[,] cellsArray, long curRow, SalesHistoryEntry entry) {
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
            entry.customerName = inputText.Substring(startIndex, endIndex - startIndex).Trim();

            // Extract customer postal code
            startIndex = inputText.LastIndexOf(',') + 1;
            entry.customerPostalCode = inputText.Substring(startIndex).Trim();
        }
        
        private static void GetItemInfo(object[,] cellsArray, long curRow, SalesHistoryEntry entry) {
            entry.orderNum = Convert.ToString(cellsArray[curRow, 3]);
            entry.invoiceNum = Convert.ToString(cellsArray[curRow, 4]);

            entry.quantity = (double)cellsArray[curRow, 15];

            entry.totalCost = (double)cellsArray[curRow, 18];
            entry.unitCost = Math.Abs(entry.totalCost) / entry.quantity;

            entry.totalPrice = (double)cellsArray[curRow, 17];
            entry.unitPrice = Math.Abs(entry.totalPrice) / entry.quantity;

            entry.marginPercent = (double)cellsArray[curRow, 10];
        }

        private static void GetInvoiceDateInfo(object[,] cellsArray, long curRow, SalesHistoryEntry entry) {
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
            entry.invoiceDate = inputText.Substring(startIndex).Trim();
        }

        private static void GetSalesLocationInfo(object[,] cellsArray, long curRow, SalesHistoryEntry entry) {
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
            entry.branchName = inputText.Substring(startIndex);

            // Then check if there are '( )' in the line, if there are update the name with what is between them
            startIndex = inputText.IndexOf('(');
            endIndex = inputText.LastIndexOf(')');

            if (startIndex != -1 && endIndex != -1 && startIndex < endIndex) {
                startIndex++;
                entry.branchName = inputText.Substring(startIndex, endIndex - startIndex).Trim();
            }

            // Extract branch num
            startIndex = inputText.IndexOf(':') + 1;
            endIndex = inputText.IndexOf('-');
            entry.branchNum = inputText.Substring(startIndex, endIndex - startIndex).Trim();
        }

        private static void GetSalesRepInfo(object[,] cellsArray, long curRow, SalesHistoryEntry entry) {
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

            // Extract the sales rep
            startIndex = inputText.LastIndexOf(':') + 1;
            entry.salesRep = inputText.Substring(startIndex).Trim();
        }

        private static void GetBranchInfo(object[,] cellsArray, long curRow, SalesHistoryEntry entry) {
            string inputText;
            int startIndex;
            int endIndex;

            // Loop back up until a cell is found that starts with "Branch : "
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
            while (!(inputText.Substring(0, inputText.Length > 10 ? 10 : inputText.Length) == "Branch  : "));

            // Extract branch name
            // First just grab everything after the '-'
            startIndex = inputText.IndexOf('-');
            entry.branchName = inputText.Substring(startIndex + 1).Trim();

            // Extract branch num
            startIndex = inputText.IndexOf(':') + 1;
            endIndex = inputText.IndexOf('-');
            entry.branchNum = inputText.Substring(startIndex, endIndex - startIndex).Trim();
        }

        private static void GetContractNumberInfo(object[,] cellsArray, long curRow, SalesHistoryEntry entry) {
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
            while (!(inputText.Substring(0, inputText.Length > 18 ? 18 : inputText.Length) == "Contract Number : "));

            // Extract the contract number
            startIndex = inputText.LastIndexOf(':') + 1;
            entry.contractNum = inputText.Substring(startIndex).Trim();
        }

        private static void GetInvoicePeriodInfo(object[,] cellsArray, long curRow, SalesHistoryEntry entry) {
            string inputText;
            int yearStartIndex;
            int periodStartIndex;

            // Loop back up until a cell is found that starts with "Period : "
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
            while (!(inputText.Substring(0, inputText.Length > 15 ? 15 : inputText.Length) == "Period : "));

            // Extract the invoice period
            // Formated "1999 - 11"
            yearStartIndex = inputText.IndexOf("Year : ") + 7;
            periodStartIndex = inputText.IndexOf("Period : ") + 9;

            entry.invoicePeriod = inputText.Substring(yearStartIndex, 4).Trim() + " - " + inputText.Substring(periodStartIndex, 2).Trim();
        }

        private static void GetProductGroupInfo(object[,] cellsArray, long curRow, SalesHistoryEntry entry) {
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
            while (!(inputText.Substring(0, inputText.Length > 17 ? 17 : inputText.Length) == "Product Group  : "));

            // Extract the contract number
            startIndex = inputText.LastIndexOf(':') + 1;
            entry.productGroup = inputText.Substring(startIndex).Trim();
        }

        private static void GetShipToInfo(object[,] cellsArray, long curRow, SalesHistoryEntry entry) {
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
            while (!(inputText.Substring(0, inputText.Length > 9 ? 9 : inputText.Length) == "ShipTo : "));

            // Extract the contract number
            startIndex = inputText.LastIndexOf(':') + 1;
            entry.customerShipToNum = inputText.Substring(startIndex).Trim();
        }

        private static void GetSupplierInfo(object[,] cellsArray, long curRow, SalesHistoryEntry entry) {
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
            while (!(inputText.Substring(0, inputText.Length > 11 ? 11 : inputText.Length) == "Supplier : "));

            // Extract the contract number
            startIndex = inputText.LastIndexOf(':') + 1;
            entry.supplier = inputText.Substring(startIndex).Trim();
        }
        
        private static void GetTakerInfo(object[,] cellsArray, long curRow, SalesHistoryEntry entry) {
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
            while (!(inputText.Substring(0, inputText.Length > 8 ? 8 : inputText.Length) == "Taker : "));

            // Extract the contract number
            startIndex = inputText.LastIndexOf(':') + 1;
            entry.taker = inputText.Substring(startIndex).Trim();
        }

        public void FormatWorksheet(Excel.Worksheet DataWs) {
            // Create new worksheet for data to be placed on
            Excel.Worksheet formattedWs = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(After: DataWs);

            // Put data into an array that will be copied in one go into the excel worksheet
            object[,] OutputArray = new object[entries.Count + 1, 26];
            
            // Loop through each entry, adding it to the array to be copied to the new excel sheet
            int index = 0;   // The index of the current entry
            foreach (SalesHistoryEntry entry in entries) {
                int curCol = 0;

                if (sortLevels.Contains("Sales Location")) {
                    if (index == 0) {
                        OutputArray[0, curCol] = "Branch Num";
                    }
                    
                    OutputArray[index + 1, curCol] = entry.branchNum;
                    curCol++;

                    if (index == 0) {
                        OutputArray[0, curCol] = "Branch Name";
                    }

                    OutputArray[index + 1, curCol] = entry.branchName;
                    curCol++;
                }

                if (sortLevels.Contains("Branch")) {
                    if (index == 0) {
                        OutputArray[0, curCol] = "Branch Num";
                    }

                    OutputArray[index + 1, curCol] = entry.branchNum;
                    curCol++;

                    if (index == 0) {
                        OutputArray[0, curCol] = "Branch Name";
                    }

                    OutputArray[index + 1, curCol] = entry.branchName;
                    curCol++;
                }

                if (sortLevels.Contains("Sales Rep")) {
                    if (index == 0) {
                        OutputArray[0, curCol] = "Sales Rep";
                    }

                    OutputArray[index + 1, curCol] = entry.salesRep;
                    curCol++;
                }

                if (sortLevels.Contains("Invoice Date")) {
                    if (index == 0) {
                        OutputArray[0, curCol] = "Invoice Date";
                    }

                    OutputArray[index + 1, curCol] = entry.invoiceDate;
                    curCol++;
                }

                if (sortLevels.Contains("Customer")) {
                    if (index == 0) {
                        OutputArray[0, curCol] = "Customer";
                    }

                    OutputArray[index + 1, curCol] = entry.customerName;
                    curCol++;

                    if (index == 0) {
                        OutputArray[0, curCol] = "Customer Postal Code";
                    }

                    OutputArray[index + 1, curCol] = entry.customerPostalCode;
                    curCol++;
                }

                if (sortLevels.Contains("Contract Number")) {
                    if (index == 0) {
                        OutputArray[0, curCol] = "Contract Number";
                    }

                    OutputArray[index + 1, curCol] = entry.contractNum;
                    curCol++;
                }

                if (sortLevels.Contains("Invoice Period")) {
                    if (index == 0) {
                        OutputArray[0, curCol] = "Invoice Period";
                    }

                    OutputArray[index + 1, curCol] = entry.invoicePeriod;
                    curCol++;
                }

                if (sortLevels.Contains("Product Group")) {
                    if (index == 0) {
                        OutputArray[0, curCol] = "Product Group";
                    }

                    OutputArray[index + 1, curCol] = entry.productGroup;
                    curCol++;
                }

                if (sortLevels.Contains("Ship To")) {
                    if (index == 0) {
                        OutputArray[0, curCol] = "Ship To";
                    }

                    OutputArray[index + 1, curCol] = entry.customerShipToNum;
                    curCol++;
                }

                if (sortLevels.Contains("Supplier")) {
                    if (index == 0) {
                        OutputArray[0, curCol] = "Supplier";
                    }

                    OutputArray[index + 1, curCol] = entry.supplier;
                    curCol++;
                }

                if (sortLevels.Contains("Taker")) {
                    if (index == 0) {
                        OutputArray[0, curCol] = "Taker";
                    }

                    OutputArray[index + 1, curCol] = entry.taker;
                    curCol++;
                }

                if (index == 0) {
                    OutputArray[0, curCol] = "Order Num";
                }

                OutputArray[index + 1, curCol] = entry.orderNum;
                curCol++;

                if (index == 0) {
                    OutputArray[0, curCol] = "Invoice Num";
                }

                OutputArray[index + 1, curCol] = entry.invoiceNum;
                curCol++;

                if (index == 0) {
                    OutputArray[0, curCol] = "Part Number";
                }

                OutputArray[index + 1, curCol] = entry.partNumber;
                curCol++;

                if (index == 0) {
                    OutputArray[0, curCol] = "Part Description";
                }

                OutputArray[index + 1, curCol] = entry.partDesc;
                curCol++;

                if (index == 0) {
                    OutputArray[0, curCol] = "Qty";
                }

                OutputArray[index + 1, curCol] = entry.quantity;
                curCol++;

                if (index == 0) {
                    OutputArray[0, curCol] = "Total Cost";
                }

                OutputArray[index + 1, curCol] = "$" + entry.totalCost.ToString("0.00");
                curCol++;

                if (index == 0) {
                    OutputArray[0, curCol] = "Unit Cost";
                }

                OutputArray[index + 1, curCol] = "$" + entry.unitCost.ToString("0.00");
                curCol++;

                if (index == 0) {
                    OutputArray[0, curCol] = "Total Price";
                }

                OutputArray[index + 1, curCol] = "$" + entry.totalPrice.ToString("0.00");
                curCol++;

                if (index == 0) {
                    OutputArray[0, curCol] = "Unit Price";
                }

                OutputArray[index + 1, curCol] = "$" + entry.unitPrice.ToString("0.00");
                curCol++;

                if (index == 0) {
                    OutputArray[0, curCol] = "Margin";
                }
                OutputArray[index + 1, curCol] = entry.marginPercent.ToString("0.0");

                index++;
            }

            // Copy the array contents onto the newly created worksheet
            formattedWs.Select();
            formattedWs.Range["A1", formattedWs.Cells[entries.Count, 26]] = OutputArray;

            // Bold the headers
            formattedWs.Range["A1:Z1"].Font.Bold = true;

            // Freeze the top row
            formattedWs.Application.ActiveWindow.SplitRow = 1;
            formattedWs.Application.ActiveWindow.FreezePanes = true;

            // Format the currency cells
            //formattedWs.Range["J2", formattedWs.Cells[entries.Count, 30]].Numberformat = "$0.00";

            // Resize columns
            formattedWs.UsedRange.Columns.AutoFit();
        }
    }
}

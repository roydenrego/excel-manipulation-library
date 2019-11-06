using Microsoft.Office.Interop.Excel;
using System;

namespace ExcelCore {

    public class Excel {
        private Application app = null;
        private Workbook currentBook = null;
        private Worksheet currentSheet = null;

        public Excel() : this(false) {

        }

        public Excel(Boolean isVisible) {
            app = new Application();
            app.Visible = isVisible;
        }

        public void Open(string path) {
            currentBook = app.Workbooks.Open(path);
            currentSheet = currentBook.Sheets[1]; // Explicit cast is not required here
        }

      

        public String GetRows() {
            int lastRow = currentSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row;

            string data = "";

            for (int index = 2; index <= lastRow; index++) {
                Array myValues = (Array)currentSheet.get_Range("A" +
                   index.ToString(), "B" + index.ToString()).Cells.Value;
                data = data + myValues.GetValue(1, 1) + " , " + myValues.GetValue(1, 2) + "\n";
            }
            return data;
        }
    }
}

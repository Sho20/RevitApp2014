using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using RevitAppIN.RevitObj;

namespace RevitAppIN.Util
{
    class WriteExlFromList
    {
        object misValue = System.Reflection.Missing.Value;
        public string filePath;
        public Excel.Application xlApp;
        public Excel.Workbook xlWorkBook;
        public Excel.Worksheet xlWorkSheet;
        List<RevitTray> TrayList;
        private double totalLength = 0;


        public WriteExlFromList(List<RevitTray> TrayList, string path)
        {
            int temp = 0;
            this.TrayList = TrayList;
            filePath = path;
            foreach (RevitTray tray in TrayList)
            {
                if (tray.type == "Cable Trays")
                {
                    totalLength += Convert.ToDouble(tray.length);
                }
            }
        }

        public void writeList()
        {
            xlApp = new Excel.Application();
            //xlWorkBook = xlApp.Workbooks.Open(filePath);
            xlWorkBook = xlApp.Workbooks.Add();
            xlWorkSheet = xlWorkBook.Worksheets.get_Item(1);
            xlWorkSheet.Select();
            string id = "A1";
            Excel.Range rng = xlWorkSheet.get_Range(id, id);
            rng.Select();
            int i = 1;


            foreach (RevitTray item in TrayList)
            {
                if (i == 1) // Report title
                {
                    rng[i, 1].Value = "Object Name";
                    rng[i, 2].Value = "Object Type";
                    rng[i, 3].Value = "Width";
                    rng[i, 4].Value = "Height";
                    rng[i, 5].Value = "Length";
                    rng[i, 6].Value = "Tray Type";
                    i++;
                }
                rng[i, 1].Value = item.objName;
                rng[i, 2].Value = item.type;
                rng[i, 3].Value = item.width;
                rng[i, 4].Value = item.height;
                if (item.type2 == "Cable Tray")
                {
                    rng[i, 5].Value = item.length + " mm";
                }
                else
                {
                    rng[i, 5].Value = item.length;
                }

                rng[i, 6].Value = item.type2;
                i++;
            }

            rng[i + 2, 5].Value = "Total Length";
            rng[i + 2, 6].Value = totalLength.ToString() + " mm";




            setShtFormat();
            xlWorkBook.SaveAs(filePath, misValue, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlNoChange, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();
            finProg();

        }
        private void setShtFormat()
        {
            Excel.Range rng = xlWorkSheet.get_Range("A1", "A6");
            rng.EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            rng[1, 1].ColumnWidth = 11.88;
            rng[1, 2].ColumnWidth = 16.88;
            rng[1, 3].ColumnWidth = 9.38;
            rng[1, 4].ColumnWidth = 9.38;
            rng[1, 5].ColumnWidth = 11.88;
            rng[1, 6].ColumnWidth = 11.88;
        }

        private void finProg() // for clear Excel applicatin completely
        {

            if (xlWorkSheet != null)
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlWorkSheet);
            }
            if (xlWorkBook != null)
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlWorkBook);
            }
            if (xlApp != null)
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlApp);
            }
        }

    }
}

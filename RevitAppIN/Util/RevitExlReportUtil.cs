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
        public Excel.Worksheet xlWorkSheet2;
        List<RevitTray> TrayList;
        private double totalLength = 0;
        private List<TrayTotalLength> TotLengList;
        private List<TrayFittingTotal> TotFitList = new List<TrayFittingTotal>();

        private List<TrayFittingTotal> subCalTotalFitting(List<RevitTray> TrayList, List<TrayFittingTotal> resList, string calType)
        {
            foreach (RevitTray item in TrayList)
            {
                if (item.type == calType)
                {
                    if (resList.Any())
                    {
                        foreach (TrayFittingTotal item2 in resList)
                        {
                            if ((item2.width == item.width) && (item2.height == item.height))
                            {
                                item2.amount += 1;
                                break;
                            }
                            else if (item2.GetHashCode() == resList.Last().GetHashCode())
                            {
                                resList.Add(new TrayFittingTotal { type = item.type2, width = item.width, height = item.height, amount = 1, description = item.description });
                                break;
                            }
                            else
                            {
                                continue;
                            }

                        }
                    }
                    else
                    {
                        resList.Add(new TrayFittingTotal { type = item.type2, width = item.width, height = item.height, amount = 1, description = item.description });
                    }

                }
            }

            return resList;
        }


        public List<TrayFittingTotal> calTotalFitting(List<RevitTray> TrayList)
        {
            List<TrayFittingTotal> resList = new List<TrayFittingTotal>();

            resList = subCalTotalFitting(TrayList, resList, "Elbow");
            //resList = subCalTotalFitting(TrayList, resList, "Tee");
            //resList = subCalTotalFitting(TrayList, resList, "Transition");
            //resList = subCalTotalFitting(TrayList, resList, "Vertical Inside Bend");
            //resList = subCalTotalFitting(TrayList, resList, "Vertical Outside Bend");


            return resList;
        }

        public List<TrayTotalLength> calTotalLength(List<RevitTray> TrayList)
        {
            List<TrayTotalLength> resList = new List<TrayTotalLength>();

            foreach (RevitTray item in TrayList)
            {
                if (item.type == "Cable Trays")
                {
                    if (resList.Any())
                    {
                        foreach (TrayTotalLength item2 in resList)
                        {
                            if ((item2.width == item.width) && (item2.height == item.height))
                            {
                                item2.totalLength += Convert.ToDouble(item.length);
                                break;
                            }
                            else if (item2.GetHashCode() == resList.Last().GetHashCode())
                            {
                                resList.Add(new TrayTotalLength { type = item.type2, width = item.width, height = item.height, totalLength = Convert.ToDouble(item.length), description = item.description, commodity = item.commodity });
                                break;
                            }
                            else
                            {
                                continue;
                            }

                        }
                    }
                    else
                    {
                        resList.Add(new TrayTotalLength { type = item.type2, width = item.width, height = item.height, totalLength = Convert.ToDouble(item.length), description = item.description, commodity = item.commodity });
                    }

                }
            }
            return resList;
        }

        public WriteExlFromList(List<RevitTray> TrayList, string path)
        {
            //int temp = 0;
            this.TrayList = TrayList;
            filePath = path;
            foreach (RevitTray tray in TrayList)
            {
                if (tray.type == "Cable Trays")
                {
                    totalLength += Convert.ToDouble(tray.length);
                }
            }


            TotLengList = calTotalLength(TrayList);
            //TotFitList = calTotalFitting(TrayList);
            var types_without_tray = from fit in TrayList
                                     where fit.type != "Cable Trays"
                                     select fit;

            var types = from fit in TrayList
                        group fit by new { fit.type2 };



            foreach (var typ in types)
            {
                var outputs = from fit in TrayList
                              where fit.type2 == typ.Key.type2
                              select fit;

                int amount = 0;
                string _type = "";
                string _width = "";
                string _height = "";
                string _description = "";
                string _commodity = "";
                if (typ.Key.type2 == "Cable Tray") continue;
                foreach (var item in outputs)
                {
                    amount++;
                    _type = item.type2;
                    _width = item.width;
                    _height = item.height;
                    _description = item.description;
                    _commodity = item.commodity;
                }

                TotFitList.Add(new TrayFittingTotal() { type = _type, width = _width, height = _height, amount = amount, description = _description, commodity = _commodity });


            }


        }

        public void writeList2() // for export total length according to different size tray.
        {

            xlWorkSheet2.Select();
            string id = "A1";
            Excel.Range rng = xlWorkSheet2.get_Range(id, id);
            rng.Select();
            int i = 1;


            foreach (TrayTotalLength tray in TotLengList)
            {
                if (i == 1) // Report title
                {
                    rng[i, 1].Value = "Type";
                    rng[i, 2].Value = "Width";
                    rng[i, 3].Value = "Height";
                    rng[i, 4].Value = "Qty";
                    rng[i, 5].Value = "Commodity Code";
                    rng[i, 6].Value = "Description";
                    i++;
                }
                rng[i, 1].Value = tray.type;
                rng[i, 2].Value = tray.width;
                rng[i, 3].Value = tray.height;
                rng[i, 4].Value = tray.totalLength + " mm";
                rng[i, 5].Value = tray.commodity;
                rng[i, 6].Value = tray.description;
                i++;
            }

            foreach (TrayFittingTotal fitting in TotFitList)
            {
                rng[i, 1].Value = fitting.type;
                rng[i, 2].Value = fitting.width;
                rng[i, 3].Value = fitting.height;
                rng[i, 4].Value = fitting.amount;
                rng[i, 5].Value = fitting.commodity;
                rng[i, 6].Value = fitting.description;

                i++;
            }


        }

        public void writeList()
        {
            xlApp = new Excel.Application();
            string temp_filePath = "D:\\Desktop\\下半年\\template.xlsx";
            xlWorkBook = xlApp.Workbooks.Open(@temp_filePath);
            //xlWorkBook = xlApp.Workbooks.Add();
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
                    rng[i, 7].Value = "Length2";
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
                rng[i, 7].Value = item.length;
                i++;
            }

            rng[i + 2, 5].Value = "Total Length";
            rng[i + 2, 6].Value = totalLength.ToString() + " mm";
            //xlWorkSheet2 = xlWorkBook.Worksheets.Add(misValue, xlWorkSheet, misValue, misValue);
            xlWorkSheet2 = xlWorkBook.Worksheets.get_Item(2);
            writeList2();



            setShtFormat();
            //setShtBorder(i, 7);
            xlWorkBook.SaveAs(filePath, misValue, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlNoChange, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();
            finProg();

        }

        private void setShtBorder(int i, int colCount)
        {
            string col = (char)(colCount + 64) + i.ToString();
            Excel.Range rng = xlWorkSheet.get_Range("A1", col);
            //rng.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, null);
            
            
            
        
        }
        private void AllBorders(Excel.Borders _borders)
        {
            _borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            _borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            _borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            _borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            
        }
        //set up the sheet format
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

            //rng.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, null);
            
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

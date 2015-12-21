using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Autodesk.Revit.UI;
using Autodesk.Revit;
using Autodesk.Revit.DB;


namespace RevitAppIN
{
    public class RevitExl
    {
        object misValue = System.Reflection.Missing.Value;
        public string filePath;
        public Excel.Application xlApp;
        public Excel.Workbook xlWorkBook;
        public Excel.Worksheet xlWorkSheet;
        public int idxCol, idxRow;
        

        
        public void iniExl()
        {
            xlApp = new Excel.Application();
            idxCol = 65;
            idxRow = 1;
           
        }

        public void oExl()
        {
            xlWorkBook = xlApp.Application.Workbooks.Open(filePath);
            
        }

        public void cExl()
        {
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();
        }

      

        public void wExl(DataTable dt)
        {
            xlWorkSheet = xlWorkBook.Worksheets.get_Item(1);
            xlWorkSheet.Select ();
           
            string id;
            id = Convert.ToChar(idxCol) + idxRow.ToString();
            Excel.Range rng = xlWorkSheet.get_Range(id, id);
            rng.Select();

            int count_rows = dt.Rows.Count;


            tbl2Exl (count_rows , rng, dt);
           


        }

        public void tbl2Exl(int i, Excel .Range rng, DataTable dt)
        {
           

            while (i > 0)
            {
                
                rng[i + 1, 1].Value = dt.Rows[i-1][0].ToString();
                rng[i + 1, 2].Value = dt.Rows[i-1][1].ToString();
                rng[i + 1, 3].Value = dt.Rows[i-1][2].ToString();
                rng[i + 1, 4].Value = dt.Rows[i-1][3].ToString();
                rng[i + 1, 5].Value = dt.Rows[i-1][4].ToString();
                rng[i + 1, 6].Value = dt.Rows[i-1][5].ToString();
                //rng[i + 1, 7].Value = dt.Rows[i-1][6].ToString();
                //rng[i + 1, 8].Value = dt.Rows[i-1][7].ToString();

                i--;
            }
            
            if (i == 0)
            {
                rng[1, 1].Value = dt.Columns[0].ColumnName.ToString();
                rng[1, 2].Value = dt.Columns[1].ColumnName.ToString();
                rng[1, 3].Value = dt.Columns[2].ColumnName.ToString();
                rng[1, 4].Value = dt.Columns[3].ColumnName.ToString();
                rng[1, 5].Value = dt.Columns[4].ColumnName.ToString();
                rng[1, 6].Value = dt.Columns[5].ColumnName.ToString();
                //rng[1, 7].Value = dt.Columns[6].ColumnName.ToString();
                //rng[1, 8].Value = dt.Columns[7].ColumnName.ToString();
            }
  
        
        }


    }
}

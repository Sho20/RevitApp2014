using System;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Autodesk.Revit;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using GXYZ = Autodesk.Revit.DB.XYZ;
using RevitAppIN.Util;
using RevitAppIN.RevitObj;

namespace RevitAppIN.Forms
{
    public partial class ReportForm : System.Windows.Forms.Form
    {
        //private ElementSet selectedObjects;
        private ArrayList selectedObjects;

        public ReportForm()
        {
            InitializeComponent();
            _myInitializeComponent();
        }
        //public ReportForm(ElementSet selection)
        public ReportForm(ArrayList selection)
        {
            InitializeComponent();
            _myInitializeComponent();
            selectedObjects = selection;

        }

        private void _myInitializeComponent()
        {
            txtInputFilePath.Text = "";

        }

        private void RunReport_Click(object sender, EventArgs e)
        {
            List<RevitTray> TrayList = new List<RevitTray>();
            foreach (Element ele in selectedObjects)
            {
                //TaskDialog.Show("test", ele.getWidth());

                if (ele == null) continue;
                
                TrayList.Add(new RevitTray() { objName = ele.getObjectID(), type = ele.getType(), type2 = ele.getType2(), width = ele.getWidth(), height = ele.getHeight(), length = ele.getLength(), description = ele.getTraySpec(),commodity = ele.getCommodity() });
                
            }

            WriteExlFromList ExlTask = new WriteExlFromList(TrayList, txtInputFilePath.Text);
            ExlTask.writeList();

            TaskDialog.Show("RevitAppIN", "Report Exported.");

        }

        private void txtInputFilePath_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                txtInputFilePath.Text = saveFileDialog1.FileName;
        }

        private void btnSetFilePath_Click(object sender, EventArgs e)
        {
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                txtInputFilePath.Text = saveFileDialog1.FileName;
        }

        
    }
}

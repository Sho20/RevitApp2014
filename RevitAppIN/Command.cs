//
// (C) Copyright 2003-2012 by Autodesk, Inc.
//
// Permission to use, copy, modify, and distribute this software in
// object code form for any purpose and without fee is hereby granted,
// provided that the above copyright notice appears in all copies and
// that both that copyright notice and the limited warranty and
// restricted rights notice below appear in all supporting
// documentation.
//
// AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS.
// AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
// MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE. AUTODESK, INC.
// DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
// UNINTERRUPTED OR ERROR FREE.
//
// Use, duplication, or disclosure by the U.S. Government is subject to
// restrictions set forth in FAR 52.227-19 (Commercial Computer
// Software - Restricted Rights) and DFAR 252.227-7013(c)(1)(ii)
// (Rights in Technical Data and Computer Software), as applicable.
//


using System;
using System.Data;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using System.Windows.Forms;
using Autodesk.Revit;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using GXYZ = Autodesk.Revit.DB.XYZ;
using RevitAppIN.Util;

//namespace Revit.SDK.Samples.ParameterUtils.CS
namespace RevitAppIN
{
    /// <summary>
    /// display a Revit element property-like form related to the selected element.
    /// </summary>
    //[Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.ReadOnly)]
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Automatic)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    [Autodesk.Revit.Attributes.Journaling(Autodesk.Revit.Attributes.JournalingMode.NoCommandData)]
    public class Command : IExternalCommand
    {
        static Document m_document;

        #region IExternalCommand Members
        /// <summary>
        /// Implement this method as an external command for Revit.
        /// </summary>
        /// <param name="commandData">An object that is passed to the external application 
        /// which contains data related to the command, 
        /// such as the application object and active view.</param>
        /// <param name="message">A message that can be set by the external application 
        /// which will be displayed if a failure or cancellation is returned by 
        /// the external command.</param>
        /// <param name="elements">A set of elements to which the external application 
        /// can add elements that are to be highlighted in case of failure or cancellation.</param>
        /// <returns>Return the status of the external command. 
        /// A result of Succeeded means that the API external method functioned as expected. 
        /// Cancelled can be used to signify that the user cancelled the external operation 
        /// at some point. Failure should be returned if the application is unable to proceed with 
        /// the operation.</returns>
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message,
            ElementSet elements)
        {
            // set out default result to failure.
            Autodesk.Revit.UI.Result retRes = Autodesk.Revit.UI.Result.Failed;

            //Autodesk.Revit.UI.UIApplication app = commandData.Application;  //before
            CollectorExts.CollectorExt.m_app = commandData.Application;

            //UIDocument revitDoc = commandData.Application.ActiveUIDocument;
            //m_document = revitDoc.Document;
            //Autodesk.Revit.DB.View view = revitDoc.Document.ActiveView;
            //ElementSet selection = commandData.Application.ActiveUIDocument.Selection.Elements;
            


            //FilteredElementCollector collector1 = new FilteredElementCollector(revitDoc.Document, view.Id);

         
            //***====================================================== ALL selection
            Autodesk.Revit.DB.Document doc = commandData.Application.ActiveUIDocument.Document;
            FilteredElementCollector elemTypeCtor = (new FilteredElementCollector(doc)).WhereElementIsElementType();
            FilteredElementCollector notElemTypeCtor = (new FilteredElementCollector(doc)).WhereElementIsNotElementType();
            FilteredElementCollector allElementCtor = elemTypeCtor.UnionWith(notElemTypeCtor);
            //ICollection<Element> ALLfounds = allElementCtor.ToElements();
            ICollection<Element> ALLfounds = elemTypeCtor.ToElements();

            ElementSet selection = commandData.Application.ActiveUIDocument.Selection.Elements;

            //***======================================================
            ArrayList objs = new ArrayList();
            foreach (Element element in selection)
            {
                if (element == null) 
                    continue;
                if (element.Category == null)
                    continue;
                if (!((element.Category.Name == "Cable Trays") || (element.Category.Name == "Cable Tray Fittings")))
                    continue;
                
                objs.Add(element);
                
            }


            //***===========================================
            //RevitAppUD.Forms.RevitAppMainForm.
            string txtForm;
            txtForm = "test";
            RevitAppIN.RevitObj.RevitObj testobj = new RevitAppIN.RevitObj.RevitObj();

            RevitAppIN.Forms.ReportForm testForm = new RevitAppIN.Forms.ReportForm(objs);
            testForm.ShowDialog();

            //RevitAppIN.Forms.RevitAppMainForm testForm = new RevitAppIN.Forms.RevitAppMainForm(testobj);
            //testForm.ShowDialog();
            //txtForm = testobj.width;
            //TaskDialog.Show("RevitAppUD", testobj.width);


            //string rvtPath = @"D:\revit\rvtEXL.xlsx";
            //RevitAppIN.RevitExl rvtExl = new RevitAppIN.RevitExl();
            //rvtExl.filePath = rvtPath;
            //rvtExl.iniExl();
            //rvtExl.oExl();

            DataTable dt = new DataTable();

            dt.Columns.Add("Type");
            dt.Columns.Add("Mark");
            dt.Columns.Add("Element ID");
            dt.Columns.Add("Connector Pt X");
            dt.Columns.Add("Connector Pt Y");
            dt.Columns.Add("Connector Pt Z");
            /*
            dt.Columns.Add("Type");
            dt.Columns.Add("Mark");
            dt.Columns.Add("Start Pt X");
            dt.Columns.Add("Start Pt Y");
            dt.Columns.Add("Start Pt Z");
            dt.Columns.Add("End Pt X");
            dt.Columns.Add("End Pt Y");
            dt.Columns.Add("End Pt Z");
            */


            string mrk, typ, ele_Id;
            string Ori_X, Ori_Y, Ori_Z;
            int count;


            /*
            string mrk, typ;
            string ScoorX, ScoorY, ScoorZ;
            string EcoorX, EcoorY, EcoorZ;
            string CcoorX, CcoorY, CcoorZ;
            DataRow eleRow = dt.NewRow();

            int count;
            */

            List<Autodesk.Revit.DB.ElementId> ids = new List<ElementId>();

            FamilyInstance[] instances = new FamilyInstance[3];
            Connector[] conns = new Connector[3];
            ConnectorSetIterator csi = null;




            //foreach (Element ele_temp in selection)
            //{
            //    //TaskDialog.Show("test",ele_temp.getWidth());
                

            //    if (ele_temp == null) continue;
            //    //TaskDialog.Show("RevitAppUD",ele_temp.Category .Name .ToString ());
            //    /*
            //    ids.Add(ele_temp.Id);
            //    int i = 0;
            //    instances[i] = ele_temp as FamilyInstance;
            //    csi = ConnectorInfo.GetConnectors(ids[i].IntegerValue).ForwardIterator();
            //    conns[i] = csi.Current as Connector;
            //    */
            //    //MEPCurve testConn = ele_temp as MEPCurve;
            //    //TaskDialog.Show("RevitAppUD", testConn.ConnectorManager.Connectors.ToString());


            //    if (ele_temp.Category.Name == "Cable Trays")
            //    {
            //        MEPCurve testConn = ele_temp as MEPCurve;
            //        foreach (Connector conn in testConn.ConnectorManager.Connectors)
            //        {
            //            //object connLocation = GetLocation(conn);
            //            object connLocation = RevitObjConnCollector.GetLocation(conn);
            //            //TaskDialog.Show("RevitAppUD", connLocation.ToString());
            //            XYZ Ori = connLocation as XYZ;
            //            typ = ele_temp.Category.Name;
            //            mrk = ele_temp.get_Parameter("Mark").AsString();
            //            ele_Id = ele_temp.Id.ToString();
            //            Ori_X = Ori.X.ToString();
            //            Ori_Y = Ori.Y.ToString();
            //            Ori_Z = Ori.Z.ToString();

            //            dt.Rows.Add();
            //            count = dt.Rows.Count;
            //            dt.Rows[count - 1][0] = typ;
            //            dt.Rows[count - 1][1] = mrk;
            //            dt.Rows[count - 1][2] = ele_Id;
            //            dt.Rows[count - 1][3] = Ori_X;
            //            dt.Rows[count - 1][4] = Ori_Y;
            //            dt.Rows[count - 1][5] = Ori_Z;

            //            //XYZ conn_Origin = conn.Origin;
            //            TaskDialog.Show("RevitAppUD", Ori_X + ", " + Ori_Y + ", " + Ori_Z);
            //        }
            //    }
            //    else if (ele_temp.Category.Name == "Cable Tray Fittings")
            //    {
            //        FamilyInstance testFam = ele_temp as FamilyInstance;
            //        MEPModel testConn = testFam.MEPModel;
            //        foreach (Connector conn in testConn.ConnectorManager.Connectors)
            //        {
            //            object connLocation = GetLocation(conn);
            //            //TaskDialog.Show("RevitAppUD", connLocation.ToString());
            //            XYZ Ori = connLocation as XYZ;
            //            typ = ele_temp.Category.Name;
            //            mrk = ele_temp.get_Parameter("Mark").AsString();
            //            ele_Id = ele_temp.Id.ToString();
            //            Ori_X = Ori.X.ToString();
            //            Ori_Y = Ori.Y.ToString();
            //            Ori_Z = Ori.Z.ToString();

            //            dt.Rows.Add();
            //            count = dt.Rows.Count;
            //            dt.Rows[count - 1][0] = typ;
            //            dt.Rows[count - 1][1] = mrk;
            //            dt.Rows[count - 1][2] = ele_Id;
            //            dt.Rows[count - 1][3] = Ori_X;
            //            dt.Rows[count - 1][4] = Ori_Y;
            //            dt.Rows[count - 1][5] = Ori_Z;

            //            //XYZ conn_Origin = conn.Origin;
            //            //TaskDialog.Show("RevitAppUD", conn_Origin.X.ToString() + ", " + conn_Origin.Y.ToString() + ", " + conn_Origin.Z.ToString());
            //        }
            //    }



            //    if (ele_temp.Category.Name == "Cable Trays" || ele_temp.Category.Name == "Cable Tray Fittings")
            //    {
            //        /*
            //        TaskDialog.Show("RevitAppUD", "OK!");
            //        ele_temp.get_Parameter("Width").SetValueString(txtForm + " mm");
            //        ele_temp.get_Parameter("Width_ud").Set(ele_temp.get_Parameter("Width").AsDouble());
            //        TaskDialog.Show("RevitAppUD", ele_temp.get_Parameter("Width_ud").AsValueString());
            //        //TaskDialog .Show ("RevitAppUD",ele_temp.Location);
            //        LocationCurve locationCurve = ele_temp.Location as LocationCurve;
            //        XYZ endPoint0 = locationCurve.Curve.get_EndPoint(0);
            //        XYZ endPoint1 = locationCurve.Curve.get_EndPoint(1);
            //        // TaskDialog.Show("RevitAppUD", endPoint0.X + ", " + endPoint0.Y + ", " + endPoint0.Z);
            //        //TaskDialog.Show("RevitAppUD", endPoint1.X + ", " + endPoint1.Y + ", " + endPoint1.Z);

            //        */


            //    }

            //    /*
            //    if (ele_temp.Category.Name == "Cable Trays")
            //    {
            //        typ = ele_temp.Category.Name;
            //        mrk = ele_temp.get_Parameter("Mark").AsString();
            //        LocationCurve locationCurve = ele_temp.Location as LocationCurve;
            //        XYZ endPoint0 = locationCurve.Curve.get_EndPoint(0);
            //        XYZ endPoint1 = locationCurve.Curve.get_EndPoint(1);
            //        ScoorX = endPoint0.X.ToString();
            //        ScoorY = endPoint0.Y.ToString();
            //        ScoorZ = endPoint0.Z.ToString();
            //        EcoorX = endPoint1.X.ToString();
            //        EcoorY = endPoint1.Y.ToString();
            //        EcoorZ = endPoint1.Z.ToString();

            //        dt.Rows.Add();
            //        count = dt.Rows.Count;
            //        dt.Rows[count-1][0] = typ;
            //        dt.Rows[count-1][1] = mrk;
            //        dt.Rows[count-1][2] = ScoorX;
            //        dt.Rows[count-1][3] = ScoorY;                         
            //        dt.Rows[count-1][4] = ScoorZ;
            //        dt.Rows[count-1][5] = EcoorX;
            //        dt.Rows[count-1][6] = EcoorY;
            //        dt.Rows[count-1][7] = EcoorZ;
 
            //    }
            //    else if (ele_temp.Category.Name == "Cable Tray Fittings")
            //    {
            //        typ = ele_temp.Category.Name;
            //        mrk = ele_temp.get_Parameter("Mark").AsString();
            //        LocationPoint locationPt = ele_temp.Location as LocationPoint;
            //        XYZ endPoint3 = locationPt.Point;
            //        CcoorX = endPoint3.X.ToString();
            //        CcoorY = endPoint3.Y.ToString();
            //        CcoorZ = endPoint3.Z.ToString();

            //        dt.Rows.Add();
            //        count = dt.Rows.Count;
            //        dt.Rows[count - 1][0] = typ;
            //        dt.Rows[count - 1][1] = mrk;
            //        dt.Rows[count - 1][2] = CcoorX;
            //        dt.Rows[count - 1][3] = CcoorY;
            //        dt.Rows[count - 1][4] = CcoorZ;
            //        dt.Rows[count - 1][5] = "";
            //        dt.Rows[count - 1][6] = "";
            //        dt.Rows[count - 1][7] = "";
                  
            //    }
            //    */


            //    /*
            //    if (ele_temp.Category.Name == "Cable Trays" || ele_temp.Category.Name == "Cable Tray Fittings")
            //    {
            //        TaskDialog.Show("RevitAppUD", "OK!");
            //        ele_temp.get_Parameter("Width").SetValueString(txtForm +" mm");
            //        ele_temp.get_Parameter("Width_ud").Set(ele_temp.get_Parameter("Width").AsDouble());
            //        TaskDialog.Show("RevitAppUD", ele_temp.get_Parameter("Width_ud").AsValueString());
            //        //TaskDialog .Show ("RevitAppUD",ele_temp.Location);
            //        LocationCurve locationCurve = ele_temp.Location as LocationCurve;
            //        XYZ endPoint0 = locationCurve.Curve.get_EndPoint(0);
            //        XYZ endPoint1 = locationCurve.Curve.get_EndPoint(1);
            //        // TaskDialog.Show("RevitAppUD", endPoint0.X + ", " + endPoint0.Y + ", " + endPoint0.Z);
            //        //TaskDialog.Show("RevitAppUD", endPoint1.X + ", " + endPoint1.Y + ", " + endPoint1.Z);

            //        FamilyInstance testFamily = ele_temp as FamilyInstance;
            //        Connector[] conns = new Connector[3];


            //    } 
            //    */
            //}

            //rvtExl.wExl(dt);
            //rvtExl.cExl();
            //TaskDialog.Show("RevitAppUD", "結束!");


            retRes = Autodesk.Revit.UI.Result.Succeeded;
            return retRes;
            //***===========================================


            // get the elements selected
            // The current selection can be retrieved from the active 
            // document via the selection object
            //SelElementSet seletion = app.ActiveUIDocument.Selection.Elements;

            /*
            //==============================================
            try
            {
                if (seletion.Size == 0)
                {
                    FilteredElementCollector collector = new FilteredElementCollector(revitDoc.Document, view.Id);
                    collector.WhereElementIsElementType();
                    FilteredElementIterator i = collector.GetElementIterator();
                    i.Reset();

                    ElementSet ss1 = commandData.Application.Application.Create.NewElementSet();
                    while (i.MoveNext())
                    {
                        Element e = i.Current as Element;
                       
                        ss1.Insert(e);
                        

                    }
                    seletion = ss1;
                }
                TaskDialog.Show("test", e.get_Parameter("Width").Definition.Name.ToString());
                retRes = Autodesk.Revit.UI.Result.Succeeded;
            }
            catch (System.Exception e)
            {
                message = e.Message;
                retRes = Autodesk.Revit.UI.Result.Failed;
            }
            return retRes;


            //==============================================
            */

            /*
            // we need to make sure that only one element is selected.
            if (seletion.Size == 0)
            {
                // we need to get the first and only element in the selection. Do this by getting 
                // an iterator. MoveNext and then get the current element.
                ElementSetIterator it = seletion.ForwardIterator();
                it.MoveNext();
                Element element = it.Current as Element;

                // Next we need to iterate through the parameters of the element,
                // as we iterating, we will store the strings that are to be displayed
                // for the parameters in a string list "parameterItems"
                List<string> parameterItems = new List<string>();
                ParameterSet parameters = element.Parameters;
                Parameter ud = element.get_Parameter("Width_ud");

                foreach (Parameter param in parameters)
                {
                    if (param == null) continue;

                    // We will make a string that has the following format,
                    // name type value
                    // create a StringBuilder object to store the string of one parameter
                    // using the character '\t' to delimit parameter name, type and value 
                    StringBuilder sb = new StringBuilder();

                    // the name of the parameter can be found from its definition.
                    sb.AppendFormat("{0}\t", param.Definition.Name);

                    //===========replace test123 from length===============
                    //=========== the default double unit is ft
                    if (param.Definition.Name == "Width")
                    {
                        //TaskDialog.Show("revit",param.AsDouble().ToString());
                        //TaskDialog.Show("revit", "AsElementID: " + param.AsElementId().ToString());
                        //TaskDialog.Show("revit", "AsDouble: " + param.AsDouble().ToString());
                        //TaskDialog.Show("revit", "AsInteger: " + param.AsInteger().ToString ());
                        //TaskDialog.Show("revit", "AsString: "+param.AsString());

                        TaskDialog.Show("revit", "UD AsValueString: " + ud.Definition.Name);
                        ud.Set(param.AsDouble());

                        //param.Set(Convert.ToDouble (0.567));
                        //RawSetParameterValue(ud, 0.567);
                        element.get_Parameter("Length_ud").Set(element.get_Parameter("Length").AsDouble());
                        Parameter Lud = element.get_Parameter("Length_ud");
                        param.SetValueString("500 mm");

                        TaskDialog.Show("revit", "UD AsValueString: " + ud.AsValueString());
                        TaskDialog.Show("revit", "UD AsDouble: " + ud.AsDouble().ToString());

                        TaskDialog.Show("revit", "LUD AsValueString: " + Lud.AsValueString());
                        TaskDialog.Show("revit", "LUD AsDouble: " + Lud.AsDouble().ToString());

                    }
                    //===========replace test123 from length===============


                    // Revit parameters can be one of 5 different internal storage types:
                    // double, int, string, Autodesk.Revit.DB.ElementId and None. 
                    // if it is double then use AsDouble to get the double value
                    // then int AsInteger, string AsString, None AsStringValue.
                    // Switch based on the storage type
                    

                    // add the completed line to the string list
                    parameterItems.Add(sb.ToString());
                }
                retRes = Autodesk.Revit.UI.Result.Succeeded;
            }
            else
            {
                message = "Please select only one element";
            }
            return retRes;
            */
        }



        //        #endregion

        /*
        public class ConnectorInfo
        {
            /// <summary>
            /// The owner's element ID
            /// </summary>
            int m_ownerId;

            /// <summary>
            /// The origin of the connector
            /// </summary>
            GXYZ m_origin;

            /// <summary>
            /// The Connector object
            /// </summary>
            Connector m_connector;

            /// <summary>
            /// The connector this object represents
            /// </summary>
            public Connector Connector
            {
                get { return m_connector; }
                set { m_connector = value; }
            }

            /// <summary>
            /// The owner ID of the connector
            /// </summary>
            public int OwnerId
            {
                get { return m_ownerId; }
                set { m_ownerId = value; }
            }

            /// <summary>
            /// The origin of the connector
            /// </summary>
            public GXYZ Origin
            {
                get { return m_origin; }
                set { m_origin = value; }
            }

            /// <summary>
            /// The constructor that finds the connector with the owner ID and origin
            /// </summary>
            /// <param name="ownerId">the ownerID of the connector</param>
            /// <param name="origin">the origin of the connector</param>
            public ConnectorInfo(int ownerId, GXYZ origin)
            {
                m_ownerId = ownerId;
                m_origin = origin;
                m_connector = ConnectorInfo.GetConnector(m_ownerId, origin);
            }

            /// <summary>
            /// The constructor that finds the connector with the owner ID and the values of the origin
            /// </summary>
            /// <param name="ownerId">the ownerID of the connector</param>
            /// <param name="x">the X value of the connector</param>
            /// <param name="y">the Y value of the connector</param>
            /// <param name="z">the Z value of the connector</param>
            public ConnectorInfo(int ownerId, double x, double y, double z)
                : this(ownerId, new GXYZ(x, y, z))
            {
            }

            /// <summary>
            /// Get the connector of the owner at the specific origin
            /// </summary>
            /// <param name="ownerId">the owner ID of the connector</param>
            /// <param name="connectorOrigin">the origin of the connector</param>
            /// <returns>if found, return the connector, or else return null</returns>
            public static Connector GetConnector(int ownerId, GXYZ connectorOrigin)
            {
                ConnectorSet connectors = GetConnectors(ownerId);
                foreach (Connector conn in connectors)
                {
                    if (conn.ConnectorType == ConnectorType.Logical)
                        continue;
                    if (conn.Origin.IsAlmostEqualTo(connectorOrigin))
                        return conn;
                }
                return null;
            }

            /// <summary>
            /// Get all the connectors of an element with a specific ID
            /// </summary>
            /// <param name="ownerId">the owner ID of the connector</param>
            /// <returns>the connector set which includes all the connectors found</returns>
            public static ConnectorSet GetConnectors(int ownerId)
            {
                Autodesk.Revit.DB.ElementId elementId = new ElementId(ownerId);
                Element element = m_document.GetElement(elementId);
                return GetConnectors(element);
            }

            /// <summary>
            /// Get all the connectors of a specific element
            /// </summary>
            /// <param name="element">the owner of the connector</param>
            /// <returns>if found, return all the connectors found, or else return null</returns>
            public static ConnectorSet GetConnectors(Autodesk.Revit.DB.Element element)
            {
                if (element == null) return null;
                FamilyInstance fi = element as FamilyInstance;
                if (fi != null && fi.MEPModel != null)
                {
                    return fi.MEPModel.ConnectorManager.Connectors;
                }
                MEPSystem system = element as MEPSystem;
                if (system != null)
                {
                    return system.ConnectorManager.Connectors;
                }

                MEPCurve duct = element as MEPCurve;
                if (duct != null)
                {
                    return duct.ConnectorManager.Connectors;
                }
                return null;
            }

            /// <summary>
            /// Find the two connectors of the specific ConnectorManager at the two locations
            /// </summary>
            /// <param name="connMgr">The ConnectorManager of the connectors to be found</param>
            /// <param name="ptn1">the location of the first connector</param>
            /// <param name="ptn2">the location of the second connector</param>
            /// <returns>The two connectors found</returns>
            public static Connector[] FindConnectors(ConnectorManager connMgr, GXYZ pnt1, GXYZ pnt2)
            {
                Connector[] result = new Connector[2];
                ConnectorSet conns = connMgr.Connectors;
                foreach (Connector conn in conns)
                {
                    if (conn.Origin.IsAlmostEqualTo(pnt1))
                    {
                        result[0] = conn;
                    }
                    else if (conn.Origin.IsAlmostEqualTo(pnt2))
                    {
                        result[1] = conn;
                    }
                }
                return result;
            }

        };
    
         
    */

        //public class LogUtility
        //{
        /// <summary>
        /// Invalid string.
        /// </summary>
        public const string InvalidString = "[!]";
        /// <summary>

        /// Get the location string of a connector
        /// </summary>
        /// <param name="conn">the connector to be read</param>
        /// <returns>the location information string of the connector</returns>
        private static object GetLocation(Connector conn)
        {

            if (conn.ConnectorType == ConnectorType.Logical)
            {
                return InvalidString;
            }
            Autodesk.Revit.DB.XYZ origin = conn.Origin;

            //return string.Format("{0},{1},{2}", origin.X, origin.Y, origin.Z);
            return origin;
        }
        //}



        //=====================================ConnectorInfo
        public class ConnectorInfo
        {
            //
            //static Document m_document;

            /// <summary>
            /// The owner's element ID
            /// </summary>
            int m_ownerId;

            /// <summary>
            /// The origin of the connector
            /// </summary>
            GXYZ m_origin;

            /// <summary>
            /// The Connector object
            /// </summary>
            Connector m_connector;


            /*
            public Document mDocument
            {
                get
                {
                    return m_document;
                }
                set
                {
                    m_document = value;
                }
        
            }
            */

            /// <summary>
            /// The connector this object represents
            /// </summary>
            public Connector Connector
            {
                get { return m_connector; }
                set { m_connector = value; }
            }

            /// <summary>
            /// The owner ID of the connector
            /// </summary>
            public int OwnerId
            {
                get { return m_ownerId; }
                set { m_ownerId = value; }
            }

            /// <summary>
            /// The origin of the connector
            /// </summary>
            public GXYZ Origin
            {
                get { return m_origin; }
                set { m_origin = value; }
            }

            /// <summary>
            /// The constructor that finds the connector with the owner ID and origin
            /// </summary>
            /// <param name="ownerId">the ownerID of the connector</param>
            /// <param name="origin">the origin of the connector</param>
            //public ConnectorInfo(Document mDocument, int ownerId, GXYZ origin)
            public ConnectorInfo(int ownerId, GXYZ origin)
            {
                //m_document = mDocument;
                m_ownerId = ownerId;
                m_origin = origin;
                m_connector = ConnectorInfo.GetConnector(m_ownerId, origin);

            }

            /// <summary>
            /// The constructor that finds the connector with the owner ID and the values of the origin
            /// </summary>
            /// <param name="ownerId">the ownerID of the connector</param>
            /// <param name="x">the X value of the connector</param>
            /// <param name="y">the Y value of the connector</param>
            /// <param name="z">the Z value of the connector</param>
            //public ConnectorInfo(Document mDocument, int ownerId, double x, double y, double z)
            //  : this(mDocument, ownerId, new GXYZ(x, y, z))
            public ConnectorInfo(int ownerId, double x, double y, double z)
                : this(ownerId, new GXYZ(x, y, z))
            {
            }

            /// <summary>
            /// Get the connector of the owner at the specific origin
            /// </summary>
            /// <param name="ownerId">the owner ID of the connector</param>
            /// <param name="connectorOrigin">the origin of the connector</param>
            /// <returns>if found, return the connector, or else return null</returns>
            public static Connector GetConnector(int ownerId, GXYZ connectorOrigin)
            {
                ConnectorSet connectors = GetConnectors(ownerId);
                foreach (Connector conn in connectors)
                {
                    if (conn.ConnectorType == ConnectorType.Logical)
                        continue;
                    if (conn.Origin.IsAlmostEqualTo(connectorOrigin))
                        return conn;
                }
                return null;
            }

            /// <summary>
            /// Get all the connectors of an element with a specific ID
            /// </summary>
            /// <param name="ownerId">the owner ID of the connector</param>
            /// <returns>the connector set which includes all the connectors found</returns>
            public static ConnectorSet GetConnectors(int ownerId)
            {
                Autodesk.Revit.DB.ElementId elementId = new ElementId(ownerId);
                Element element = m_document.GetElement(elementId);
                return GetConnectors(element);
            }

            /// <summary>
            /// Get all the connectors of a specific element
            /// </summary>
            /// <param name="element">the owner of the connector</param>
            /// <returns>if found, return all the connectors found, or else return null</returns>
            public static ConnectorSet GetConnectors(Autodesk.Revit.DB.Element element)
            {
                if (element == null) return null;
                FamilyInstance fi = element as FamilyInstance;
                if (fi != null && fi.MEPModel != null)
                {
                    return fi.MEPModel.ConnectorManager.Connectors;
                }
                MEPSystem system = element as MEPSystem;
                if (system != null)
                {
                    return system.ConnectorManager.Connectors;
                }

                MEPCurve duct = element as MEPCurve;
                if (duct != null)
                {
                    return duct.ConnectorManager.Connectors;
                }
                return null;
            }

            /// <summary>
            /// Find the two connectors of the specific ConnectorManager at the two locations
            /// </summary>
            /// <param name="connMgr">The ConnectorManager of the connectors to be found</param>
            /// <param name="ptn1">the location of the first connector</param>
            /// <param name="ptn2">the location of the second connector</param>
            /// <returns>The two connectors found</returns>
            public static Connector[] FindConnectors(ConnectorManager connMgr, GXYZ pnt1, GXYZ pnt2)
            {
                Connector[] result = new Connector[2];
                ConnectorSet conns = connMgr.Connectors;
                foreach (Connector conn in conns)
                {
                    if (conn.Origin.IsAlmostEqualTo(pnt1))
                    {
                        result[0] = conn;
                    }
                    else if (conn.Origin.IsAlmostEqualTo(pnt2))
                    {
                        result[1] = conn;
                    }
                }
                return result;
            }



        }
        //=====================================ConnectorInfo

    }

    /*    
        public class paraReplace(Element element, object Selection,  )
        {   
               // we need to make sure that only one element is selected.
            
                    // we need to get the first and only element in the selection. Do this by getting 
                    // an iterator. MoveNext and then get the current element.
          

                    // Next we need to iterate through the parameters of the element,
                    // as we iterating, we will store the strings that are to be displayed
                    // for the parameters in a string list "parameterItems"
                    List<string> parameterItems = new List<string>();
                    ParameterSet parameters = element.Parameters;
                    Parameter ud = element.get_Parameter("Width_ud");

                    foreach (Parameter param in parameters)
                    {
                        if (param == null) continue;

                        // We will make a string that has the following format,
                        // name type value
                        // create a StringBuilder object to store the string of one parameter
                        // using the character '\t' to delimit parameter name, type and value 
                        StringBuilder sb = new StringBuilder();

                        // the name of the parameter can be found from its definition.
                        sb.AppendFormat("{0}\t", param.Definition.Name);

                        //===========replace test123 from length===============
                        //=========== the default double unit is ft
                        if (param.Definition.Name == "Width")
                        {
                            //TaskDialog.Show("revit",param.AsDouble().ToString());
                            //TaskDialog.Show("revit", "AsElementID: " + param.AsElementId().ToString());
                            //TaskDialog.Show("revit", "AsDouble: " + param.AsDouble().ToString());
                            //TaskDialog.Show("revit", "AsInteger: " + param.AsInteger().ToString ());
                            //TaskDialog.Show("revit", "AsString: "+param.AsString());

                            TaskDialog.Show("revit", "UD AsValueString: " + ud.Definition.Name);
                            ud.Set(param.AsDouble());

                            //param.Set(Convert.ToDouble (0.567));
                            //RawSetParameterValue(ud, 0.567);
                            element.get_Parameter("Length_ud").Set(element.get_Parameter("Length").AsDouble());
                            Parameter Lud = element.get_Parameter("Length_ud");
                            param.SetValueString("500 mm");

                            TaskDialog.Show("revit", "UD AsValueString: " + ud.AsValueString());
                            TaskDialog.Show("revit", "UD AsDouble: " + ud.AsDouble().ToString());

                            TaskDialog.Show("revit", "LUD AsValueString: " + Lud.AsValueString());
                            TaskDialog.Show("revit", "LUD AsDouble: " + Lud.AsDouble().ToString());

                        }
                        //===========replace test123 from length===============


                        // Revit parameters can be one of 5 different internal storage types:
                        // double, int, string, Autodesk.Revit.DB.ElementId and None. 
                        // if it is double then use AsDouble to get the double value
                        // then int AsInteger, string AsString, None AsStringValue.
                        // Switch based on the storage type
                    

                        // add the completed line to the string list
                        parameterItems.Add(sb.ToString());
                    }
                    retRes = Autodesk.Revit.UI.Result.Succeeded;        
        }
    */
        #endregion



}

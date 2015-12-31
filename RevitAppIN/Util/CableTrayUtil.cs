using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using LocXYZ = Autodesk.Revit.DB.XYZ;
using RevitAppIN.RevitObj;

namespace RevitAppIN.Util
{
    public static class CableTrayUtil
    {
        public static string getObjectID(this Element cObject)
        {
            return cObject.Id.ToString();
        }

        public static string getType(this Element cObject)
        {
            return cObject.Category.Name;
        }

        public static string getType2(this Element cObject)
        {
            string returnValue = "Cable Tray";
            if (cObject.Category.Name == "Cable Tray Fittings")
            {
                FamilyInstance obj2 = cObject as FamilyInstance;
                returnValue = transType(obj2.Symbol.Family.Name);
            }
            return returnValue;
        }

        public static string getWidth(this Element cObject)
        {
            string returnValue = "null"; // default value is null.

            if (cObject.Category.Name == "Cable Trays")
            {
                returnValue = cObject.get_Parameter("Width").AsValueString();
            }
            else if (cObject.Category.Name == "Cable Tray Fittings")
            {
                if (cObject.get_Parameter("Tray Width") != null)
                    returnValue = cObject.get_Parameter("Tray Width").AsValueString();
                else if (cObject.get_Parameter("Tray Width 1").HasValue)
                    returnValue = cObject.get_Parameter("Tray Width 1").AsValueString();
            }

            return returnValue;

        }

        public static string getLength(this Element cObject)
        {

            string returnValue = "null"; // default value is null.

            if (cObject.Category.Name == "Cable Trays")
            {
                returnValue = cObject.get_Parameter("Length").AsValueString();
            }
            else if (cObject.Category.Name == "Cable Tray Fittings")
            {
                returnValue = cObject.get_Parameter("Tray Length").AsValueString();
            }
            return returnValue;
        }


        public static string getHeight(this Element cObject)
        {
            try
            {
                string returnValue = "null"; // default value is null.

                if (cObject.Category.Name == "Cable Trays")
                {
                    returnValue = cObject.get_Parameter("Height").AsValueString();
                }
                else if (cObject.Category.Name == "Cable Tray Fittings")
                {
                    FamilyInstance obj2 = cObject as FamilyInstance;
                    if (transType(obj2.Symbol.Family.Name) == "Transition")
                    {
                        returnValue = cObject.get_Parameter("Tray Height1").AsValueString();
                    }
                    else
                    {
                        returnValue = cObject.get_Parameter("Tray Height").AsValueString();
                    }

                }
                return returnValue;
            }
            catch
            {
                FamilyInstance obj2 = cObject as FamilyInstance;
                //TaskDialog.Show("RevitAppIN", obj2.Symbol.Family.Name);
                return "null";
            }





        }

        private static string transType(string revitType)
        {
            string fittingType = "null";
            switch (revitType)
            {
                case "M_Channel Horizontal Bend":
                    fittingType = "Elbow";
                    break;
                case "M_Channel Vertical Inside Bend":
                    fittingType = "Vertical Inside Bend";
                    break;
                case "M_Channel Vertical Outside Bend":
                    fittingType = "Vertical Outside Bend";
                    break;
                case "M_Channel Horizontal Tee":
                    fittingType = "Tee";
                    break;
                case "M_Channel Horizontal Cross":
                    fittingType = "Cross";
                    break;
                case "M_Channel Reducer":
                    fittingType = "Transition";
                    break;
                case "M_Channel Union":
                    fittingType = "Union";
                    break;
            }
            return fittingType;
        }





        public static string getCommodity(this Element cObject)
        {
            string commodity = "null";

            if (cObject.getType2() == "Cable Tray")
            {
                if (cObject.getWidth() == "300 mm" && cObject.getHeight() == "150 mm")
                    commodity = "XEKAAAJ0-M300";
                if (cObject.getWidth() == "600 mm" && cObject.getHeight() == "150 mm")
                    commodity = "XEKAAAJ0-M600";
                if (cObject.getWidth() == "900 mm" && cObject.getHeight() == "150 mm")
                    commodity = "XEKAAAJ0-M900";
            }
            if (cObject.getType2() == "Elbow")
            {
                if (cObject.getWidth() == "300 mm" && cObject.getHeight() == "150 mm")
                    commodity = "XELBAAJC40-M300";
                if (cObject.getWidth() == "600 mm" && cObject.getHeight() == "150 mm")
                    commodity = "XELBAAJC40-M600";
                if (cObject.getWidth() == "900 mm" && cObject.getHeight() == "150 mm")
                    commodity = "XELBAAJC40-M900";
            }
            if (cObject.getType2() == "Vertical Inside Bend")
            {
                if (cObject.getWidth() == "300 mm" && cObject.getHeight() == "150 mm")
                    commodity = "XELEAAJC40-M300";
                if (cObject.getWidth() == "600 mm" && cObject.getHeight() == "150 mm")
                    commodity = "XELEAAJC40-M600";
                if (cObject.getWidth() == "900 mm" && cObject.getHeight() == "150 mm")
                    commodity = "XELEAAJC40-M900";
            }
            if (cObject.getType2() == "Vertical Outside Bend")
            {
                if (cObject.getWidth() == "300 mm" && cObject.getHeight() == "150 mm")
                    commodity = "XELFAAJC40-M300";
                if (cObject.getWidth() == "600 mm" && cObject.getHeight() == "150 mm")
                    commodity = "XELFAAJC40-M600";
                if (cObject.getWidth() == "900 mm" && cObject.getHeight() == "150 mm")
                    commodity = "XELFAAJC40-M900";
            }
            if (cObject.getType2() == "Tee")
            {
                if (cObject.getWidth() == "300 mm" && cObject.getHeight() == "150 mm")
                    commodity = "XELCAAJC40-M300";
                if (cObject.getWidth() == "600 mm" && cObject.getHeight() == "150 mm")
                    commodity = "XELCAAJC40-M600";
                if (cObject.getWidth() == "900 mm" && cObject.getHeight() == "150 mm")
                    commodity = "XELCAAJC40-M900";
            }
            if (cObject.getType2() == "Cross")
            {
                if (cObject.getWidth() == "300 mm" && cObject.getHeight() == "150 mm")
                    commodity = "XELDAAJC40-M300";
                if (cObject.getWidth() == "600 mm" && cObject.getHeight() == "150 mm")
                    commodity = "XELDAAJC40-M600";
                if (cObject.getWidth() == "900 mm" && cObject.getHeight() == "150 mm")
                    commodity = "XELDAAJC40-M900";
            }
            if (cObject.getType2() == "Transition")
            {
                if (cObject.getWidth() == "300 mm" && cObject.getHeight() == "150 mm")
                    commodity = "XELHAAJC40-M300";
                if (cObject.getWidth() == "600 mm" && cObject.getHeight() == "150 mm")
                    commodity = "XELHAAJC40-M600";
                if (cObject.getWidth() == "900 mm" && cObject.getHeight() == "150 mm")
                    commodity = "XELHAAJC40-M900";
            }


            return commodity;
        }

        //// another method to get the part description.
        //public static string getTrayDescription(string commodity)
        //{
        //    string returnValue = "null";

        //    return returnValue;
        //}

        // Get the description by using this mehtod for now, we should build a database to search from for furture work.
        public static string getTraySpec(this Element cObject)
        {
            string traySpec = "null";
            if (cObject.getType2() == "Cable Tray")
            {
                if (cObject.getWidth() == "600 mm" && cObject.getHeight() == "150 mm")
                    traySpec = "PUNCH TYPE CABLE TRAY, STRAIGHT WAY, ALUMINUM ALLOY, ANODIZING, 150MM SIDE RAIL, 600 MM";
                if (cObject.getWidth() == "900 mm" && cObject.getHeight() == "150 mm")
                    traySpec = "PUNCH TYPE CABLE TRAY, STRAIGHT WAY, ALUMINUM ALLOY, ANODIZING, 150MM SIDE RAIL, 900 MM";
                if (cObject.getWidth() == "300 mm" && cObject.getHeight() == "150 mm")
                    traySpec = "PUNCH TYPE CABLE TRAY, STRAIGHT WAY, ALUMINUM ALLOY, ANODIZING, 150MM SIDE RAIL, 300 MM";
            }
            if (cObject.getType2() == "Elbow")
            {
                if (cObject.getWidth() == "300 mm" && cObject.getHeight() == "150 mm")
                    traySpec = "PUNCH TYPE CABLE TRAY FITTING, HORIZONTAL BEND, ALUMINUM ALLOY, ANODIZING, 100MM SIDE RAIL, 300MM BEND RADIUS, 90 DEGREE BEND, 300 MM";
                if (cObject.getWidth() == "600 mm" && cObject.getHeight() == "150 mm")
                    traySpec = "PUNCH TYPE CABLE TRAY FITTING, HORIZONTAL BEND, ALUMINUM ALLOY, ANODIZING, 100MM SIDE RAIL, 300MM BEND RADIUS, 90 DEGREE BEND, 600 MM";
                if (cObject.getWidth() == "900 mm" && cObject.getHeight() == "150 mm")
                    traySpec = "PUNCH TYPE CABLE TRAY FITTING, HORIZONTAL BEND, ALUMINUM ALLOY, ANODIZING, 100MM SIDE RAIL, 300MM BEND RADIUS, 90 DEGREE BEND, 900 MM";
            }
            if (cObject.getType2() == "Vertical Inside Bend")
            {
                if (cObject.getWidth() == "300 mm" && cObject.getHeight() == "150 mm")
                    traySpec = "PUNCH TYPE CABLE TRAY FITTING, VERTICAL INSIDE BEND, ALUMINUM ALLOY, ANODIZING, 100MM SIDE RAIL, 300MM BEND RADIUS, 90 DEGREE BEND, 300 MM";
                if (cObject.getWidth() == "600 mm" && cObject.getHeight() == "150 mm")
                    traySpec = "PUNCH TYPE CABLE TRAY FITTING, VERTICAL INSIDE BEND, ALUMINUM ALLOY, ANODIZING, 100MM SIDE RAIL, 300MM BEND RADIUS, 90 DEGREE BEND, 600 MM";
                if (cObject.getWidth() == "900 mm" && cObject.getHeight() == "150 mm")
                    traySpec = "PUNCH TYPE CABLE TRAY FITTING, VERTICAL INSIDE BEND, ALUMINUM ALLOY, ANODIZING, 100MM SIDE RAIL, 300MM BEND RADIUS, 90 DEGREE BEND, 900 MM";
            }
            if (cObject.getType2() == "Vertical Outside Bend")
            {
                if (cObject.getWidth() == "300 mm" && cObject.getHeight() == "150 mm")
                    traySpec = "PUNCH TYPE CABLE TRAY FITTING, VERTICAL OUTSIDE BEND, ALUMINUM ALLOY, ANODIZING, 100MM SIDE RAIL, 300MM BEND RADIUS, 45 DEGREE BEND, 300 MM";
                if (cObject.getWidth() == "600 mm" && cObject.getHeight() == "150 mm")
                    traySpec = "PUNCH TYPE CABLE TRAY FITTING, VERTICAL OUTSIDE BEND, ALUMINUM ALLOY, ANODIZING, 100MM SIDE RAIL, 300MM BEND RADIUS, 45 DEGREE BEND, 600 MM";
                if (cObject.getWidth() == "900 mm" && cObject.getHeight() == "150 mm")
                    traySpec = "PUNCH TYPE CABLE TRAY FITTING, VERTICAL OUTSIDE BEND, ALUMINUM ALLOY, ANODIZING, 100MM SIDE RAIL, 300MM BEND RADIUS, 45 DEGREE BEND, 900 MM";
            }
            if (cObject.getType2() == "Tee")
            {
                if (cObject.getWidth() == "300 mm" && cObject.getHeight() == "150 mm")
                    traySpec = "PUNCH TYPE CABLE TRAY FITTING, HORIZONTAL TEE, ALUMINUM ALLOY, ANODIZING, 100MM SIDE RAIL, 300MM BEND";
                if (cObject.getWidth() == "600 mm" && cObject.getHeight() == "150 mm")
                    traySpec = "PUNCH TYPE CABLE TRAY FITTING, HORIZONTAL TEE, ALUMINUM ALLOY, ANODIZING, 100MM SIDE RAIL, 600MM BEND";
                if (cObject.getWidth() == "900 mm" && cObject.getHeight() == "150 mm")
                    traySpec = "PUNCH TYPE CABLE TRAY FITTING, HORIZONTAL TEE, ALUMINUM ALLOY, ANODIZING, 100MM SIDE RAIL, 900MM BEND";
            }
            if (cObject.getType2() == "Cross")
            {
                if (cObject.getWidth() == "300 mm" && cObject.getHeight() == "150 mm")
                    traySpec = "PUNCH TYPE CABLE TRAY FITTING, HORIZONTAL CROSS, STEEL SHEET, HOT-DIP GALVANIZED, 120MM SIDE RAIL, 300MM BEND RADIUS, 300 MM";
                if (cObject.getWidth() == "600 mm" && cObject.getHeight() == "150 mm")
                    traySpec = "PUNCH TYPE CABLE TRAY FITTING, HORIZONTAL CROSS, STEEL SHEET, HOT-DIP GALVANIZED, 120MM SIDE RAIL, 300MM BEND RADIUS, 600 MM";
                if (cObject.getWidth() == "900 mm" && cObject.getHeight() == "150 mm")
                    traySpec = "PUNCH TYPE CABLE TRAY FITTING, HORIZONTAL CROSS, STEEL SHEET, HOT-DIP GALVANIZED, 120MM SIDE RAIL, 300MM BEND RADIUS, 900 MM";
            }
            if (cObject.getType2() == "Transition")
            {
                if (cObject.getWidth() == "300 mm" && cObject.getHeight() == "150 mm")
                    traySpec = "PUNCH TYPE CABLE TRAY FITTING, RIGHT REDUCER, STEEL SHEET, HOT-DIP GALVANIZED, 120MM SIDE RAIL, 600X300 MM";
                if (cObject.getWidth() == "600 mm" && cObject.getHeight() == "150 mm")
                    traySpec = "PUNCH TYPE CABLE TRAY FITTING, RIGHT REDUCER, STEEL SHEET, HOT-DIP GALVANIZED, 120MM SIDE RAIL, 900X600 MM";
                if (cObject.getWidth() == "900 mm" && cObject.getHeight() == "150 mm")
                    traySpec = "PUNCH TYPE CABLE TRAY FITTING, RIGHT REDUCER, STEEL SHEET, HOT-DIP GALVANIZED, 120MM SIDE RAIL, 900X900 MM";
            }



            return traySpec;
        }

        // ====== it shall be changed in each case. =======
        //private static string transType(string revitType)
        //{
        //    string fittingType = "null";
        //    switch (revitType)
        //    {
        //        case "M_Ladder Horizontal Bend":
        //            fittingType = "Elbow";
        //            break;
        //        case "M_Ladder Horizontal Tee":
        //            fittingType = "Tee";
        //            break;
        //        case "M_Ladder Vertical Outside Bend":
        //            fittingType = "Outside Bend";
        //            break;

        //        default:
        //            fittingType = "Cable Tray";
        //            break;
        //    }

        //    return fittingType;
        //}

        /*
        public static string getLocPt(this Element cObject)
        {
            string returnValue = "null";
            if (cObject.Category.Name == "Cable Trays")
            {
                LocXYZ pt = cObject.Location


                returnValue = cObject.get_Parameter("Height").AsValueString();
            }

            return returnValue;
        }
        */
    }


}

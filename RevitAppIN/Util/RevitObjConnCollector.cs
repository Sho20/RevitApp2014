using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.Revit.DB;

namespace RevitAppIN.Util
{
    class RevitObjConnCollector
    {
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
        public static object GetLocation(Connector conn)
        {

            if (conn.ConnectorType == ConnectorType.Logical)
            {
                return InvalidString;
            }
            Autodesk.Revit.DB.XYZ origin = conn.Origin;

            //return string.Format("{0},{1},{2}", origin.X, origin.Y, origin.Z);
            return origin;
        }
    }
}

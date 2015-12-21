using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RevitAppIN.RevitObj
{
    class Route
    {
        private string _routeName;
        public string RouteName
        {
            get {return _routeName;}
            set {_routeName = value;}
        }

        private List<RouteClass> RouteList;


        void AddItem(string objName, string locX, string locY, string locZ) 
        {
            
        }

    }

    public class RouteClass
    {
        public string objName { get; set; }
        public string locX { get; set; }
        public string locY { get; set; }
        public string locZ { get; set; }
    
    }
}

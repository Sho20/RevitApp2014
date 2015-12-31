using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RevitAppIN.RevitObj
{
    public class RevitObj
    {

        public string width;

        public string getWidth()
        {
            return width;
        }


    }

    public class RevitTray
    {
        // Cable Tray and Cable Tray Fitting
        public string objName { get; set; }
        public string type { get; set; }
        public string type2 { get; set; }
        public string width { get; set; }
        public string height { get; set; }
        public string length { get; set; }
        public string description { get; set; }
        public string commodity { get; set; }

        public List<string> getInfo()
        {
            List<string> returnList = new List<string>() { type2, width, height, length, description};
            return returnList;
        }
    }

    public class TrayTotalLength
    {
        public string type { get; set; }
        public string width { get; set; }
        public string height { get; set; }
        public double totalLength { get; set; }
        public string description { get; set; }
        public string commodity { get; set; }
    }
    public class TrayFittingTotal
    {
        public string type { get; set; }
        public string width { get; set; }
        public string height { get; set; }
        public int amount { get; set; }
        public string description { get; set; }
        public string commodity { get; set; }

    }


    public class TraySpec
    {
        public string width { get; set; }
        public string height { get; set; }
        public string description { get; set; }
    }

    
}

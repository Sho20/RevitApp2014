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
    }

    public class TrayTotalLength
    {
        public string width { get; set; }
        public string height { get; set; }
        public double totalLength { get; set; }
    }
}

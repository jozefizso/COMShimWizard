using System;
using System.Collections.Generic;
using System.Text;
using Office = Microsoft.Office.Core;

namespace TestOutlook2007Addin
{
    public class InspectorCloseHandler
    {
        private Connect addin;
        private Office.CustomTaskPane ctp;

        public InspectorCloseHandler(Connect addin, Office.CustomTaskPane ctp)
        {
            this.addin = addin;
            this.ctp = ctp;
        }

        // When the Inspector closes, remove its CTP.
        public void CloseEventHandler(ref bool Cancel)
        {
            addin.inspectorPanes.Remove(ctp);
            ctp.Delete();
        }
    }
}

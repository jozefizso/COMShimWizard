namespace TestInfoPathAddin
{
    using System;
    using Extensibility;
    using System.Runtime.InteropServices;
    using InfoPath = Microsoft.Office.Interop.InfoPath;
    using System.Windows.Forms;


    [GuidAttribute("5699ACA4-EE8F-4FD3-8D4F-05078C3FD6A5"), ProgId("TestInfoPathAddin.Connect")]
    public class Connect : Object, Extensibility.IDTExtensibility2
    {
        public Connect()
        {
        }

        public void OnConnection(object application, 
            Extensibility.ext_ConnectMode connectMode, object addInInst, ref System.Array custom)
        {
            infoPathApp = (InfoPath.ApplicationClass)application;
            addInInstance = addInInst;
        }

        public void OnDisconnection(
            Extensibility.ext_DisconnectMode disconnectMode, ref System.Array custom)
        {
        }

        public void OnAddInsUpdate(ref System.Array custom)
        {
        }

        public void OnStartupComplete(ref System.Array custom)
        {
            MessageBox.Show(infoPathApp.Name);
        }

        public void OnBeginShutdown(ref System.Array custom)
        {
        }

        private InfoPath.ApplicationClass infoPathApp;
        private object addInInstance;
    }
}
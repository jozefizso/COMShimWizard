using System;
using Extensibility;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Vbe.Interop.Forms;
using System.Drawing;
using System.Collections.Generic;
using Microsoft.Win32;
using System.Configuration;

namespace TestOutlook2007Addin
{
    [Guid(Connect.TypeGuid)]
    [ProgId(Connect.TypeProgId)]
    [ComVisible(true)]
    public class Connect :
        Extensibility.IDTExtensibility2,
        Office.ICustomTaskPaneConsumer,
        Office.IRibbonExtensibility,
        Outlook.FormRegionStartup
    {

        #region fields

        private string taskPaneTitle;
        private Outlook.Application olkApp;
        private Outlook.Inspectors inspectors;
        private Outlook.ItemEvents_10_Event itemEvents;
        internal Dictionary<Office.CustomTaskPane, Outlook.Inspector> inspectorPanes;
        private Office.ICTPFactory taskPaneFactory;

        public const string TypeGuid = "BF37AC71-BB7E-43E6-BCD2-DA071449E9A1";
        public const string TypeProgId = "TestOutlook2007Addin.Connect";
        internal const string OfficeAddinKeyName =
            @"Software\Microsoft\Office\Outlook\Addins\" + Connect.TypeProgId;
        internal const string FormRegionTaskKeyName =
            @"Software\Microsoft\Office\Outlook\FormRegions\IPM.Task\";
        internal const string FormRegionNoteKeyName =
            @"Software\Microsoft\Office\Outlook\FormRegions\IPM.Note\";
        internal const string FormRegionValue1 = "CustomFormRegion_1";
        internal const string FormRegionValue2 = "CustomFormRegion_2";
        internal const string FormRegionValue3 = "CustomFormRegion_3";

        // We allow for multiple instances of the custom form region, so we need
        // somewhere to cache these.
        private List<FormRegionControls> openRegions = new List<FormRegionControls>();

        #endregion


        #region Initialization

        public Connect()
        {
            taskPaneTitle = ConfigurationManager.AppSettings["taskPaneTitle"];
            if (taskPaneTitle == null)
            {
                taskPaneTitle = "default";
            }
        }

        #endregion


        #region IDTExtensibility2

        public void OnConnection(
            object application, Extensibility.ext_ConnectMode connectMode,
            object addInInst, ref System.Array custom)
        {
            olkApp = (Outlook.Application)application;

            // Setup a mapping between each Inspector and its task pane.
            inspectorPanes = 
                new Dictionary<Office.CustomTaskPane, Outlook.Inspector>();

            // Sink the NewInspector events.
            inspectors = this.olkApp.Inspectors;
            inspectors.NewInspector +=
                new Outlook.InspectorsEvents_NewInspectorEventHandler(
                inspectors_NewInspector);
        }

        public void OnDisconnection(
            Extensibility.ext_DisconnectMode disconnectMode,
            ref System.Array custom)
        {
        }

        public void OnAddInsUpdate(ref System.Array custom)
        {
        }

        public void OnStartupComplete(ref System.Array custom)
        {
        }

        public void OnBeginShutdown(ref System.Array custom)
        {
        }

        #endregion


        #region Registration

        // We only need to explicitly register these subkeys:

        // [HKCU\Software\Microsoft\Office\Outlook\Addins\TestOutlook2007Addin.Connect]
        // "FriendlyName"="TestOutlook2007Addin"
        // "Description"=""
        // "LoadBehavior"=dword:00000003
        //
        //[HKCU\Software\Microsoft\Office\Outlook\FormRegions\IPM.Note]
        //"CustomFormRegion_1"="=TestOutlook2007Addin.Connect"
        //"CustomFormRegion_2"="=TestOutlook2007Addin.Connect"
        //
        //[HKCU\Software\Microsoft\Office\Outlook\FormRegions\IPM.Task]
        //"CustomFormRegion_1"="=TestOutlook2007Addin.Connect"
        //"CustomFormRegion_3"="=TestOutlook2007Addin.Connect"
        //
        // All other COM keys are created because we specify Register for
        // COM interop in the project properties.

        [ComRegisterFunction]
        public static void RegisterFunction(Type type)
        {
            using (RegistryKey officeAddinKey =
                Registry.CurrentUser.CreateSubKey(Connect.OfficeAddinKeyName))
            {
                officeAddinKey.SetValue("FriendlyName", Connect.TypeProgId);
                officeAddinKey.SetValue("Description", "");
                officeAddinKey.SetValue("LoadBehavior", 3);
            }

            using (RegistryKey formRegionKey =
                Registry.CurrentUser.CreateSubKey(Connect.FormRegionNoteKeyName))
            {
                formRegionKey.SetValue(
                    Connect.FormRegionValue1, "=" + Connect.TypeProgId);
                formRegionKey.SetValue(
                    Connect.FormRegionValue2, "=" + Connect.TypeProgId);
            }
            using (RegistryKey formRegionKey =
                Registry.CurrentUser.CreateSubKey(Connect.FormRegionTaskKeyName))
            {
                formRegionKey.SetValue(
                    Connect.FormRegionValue1, "=" + Connect.TypeProgId);
                formRegionKey.SetValue(
                    Connect.FormRegionValue3, "=" + Connect.TypeProgId);
            }
        }

        [ComUnregisterFunction]
        public static void UnRegisterFunction(Type type)
        {
            Registry.CurrentUser.DeleteSubKey(Connect.OfficeAddinKeyName);

            using (RegistryKey formRegionKey =
                Registry.CurrentUser.CreateSubKey(Connect.FormRegionNoteKeyName))
            {
                formRegionKey.DeleteValue(Connect.FormRegionValue1);
                formRegionKey.DeleteValue(Connect.FormRegionValue2);
            }
            using (RegistryKey formRegionKey =
                Registry.CurrentUser.CreateSubKey(Connect.FormRegionTaskKeyName))
            {
                formRegionKey.DeleteValue(Connect.FormRegionValue1);
                formRegionKey.DeleteValue(Connect.FormRegionValue3);
            }
        }

        #endregion


        #region IRibbonExtensibility

        public string GetCustomUI(string ribbonID)
        {
            return Resources.RibbonX;
        }

        public void OnToggleTaskPane(
            Office.IRibbonControl control, bool isPressed)
        {
            // Find the task pane that maps to the active Inspector, and 
            // toggle its visibility.
            Outlook.Inspector inspector = olkApp.ActiveInspector();
            foreach (KeyValuePair<Office.CustomTaskPane,
                Outlook.Inspector> keypair in inspectorPanes)
            {
                if (keypair.Value == inspector)
                {
                    Office.CustomTaskPane taskPane = keypair.Key;
                    taskPane.Visible = isPressed;
                    break;
                }
            }
        }

        #endregion


        #region ICustomTaskPaneConsumer

        public void CTPFactoryAvailable(Office.ICTPFactory CTPFactoryInst)
        {
            taskPaneFactory = CTPFactoryInst;
        }

        void inspectors_NewInspector(Outlook.Inspector Inspector)
        {
            try
            {
                // When a new Inspector opens, create a task pane and attach
                // it to this Inspector. Also add the task pane<-->Inspector 
                // mapping to the collection.
                Office.CustomTaskPane taskPane = taskPaneFactory.CreateCTP(
                    "TestOutlook2007Addin.SimpleControl", taskPaneTitle,
                    Inspector);
                inspectorPanes.Add(taskPane, Inspector);

                // Sink the Close event on this Inspector to make sure the 
                // task pane is also destroyed.
                itemEvents =
                    (Outlook.ItemEvents_10_Event)Inspector.CurrentItem;
                InspectorCloseHandler chc = 
                    new InspectorCloseHandler(this, taskPane);
                itemEvents.Close += 
                    new Outlook.ItemEvents_10_CloseEventHandler(
                    chc.CloseEventHandler);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        #endregion


        #region FormRegionStartup

        public object GetFormRegionStorage(
            string FormRegionName, object Item, int LCID,
            Outlook.OlFormRegionMode FormRegionMode,
            Outlook.OlFormRegionSize FormRegionSize)
        {
            Application.DoEvents();
            switch (FormRegionName)
            {
                case Connect.FormRegionValue1:
                    return Resources.CustomFormRegion_1_OFS;
                case Connect.FormRegionValue2:
                    return Resources.CustomFormRegion_2_OFS;
                case Connect.FormRegionValue3:
                    return Resources.CustomFormRegion_3_OFS;
                default:
                    return null;
            }
        }

        public void BeforeFormRegionShow(Outlook.FormRegion FormRegion)
        {
            // Create a new wrapper for the form region controls, hook up the Closed
            // event, and add it to our collection.
            FormRegionControls regionControls = new FormRegionControls(FormRegion);
            regionControls.Close += new EventHandler(regionControls_Close);
            openRegions.Add(regionControls);
        }

        void regionControls_Close(object sender, EventArgs e)
        {
            // When the user closes this form region, we remove the controls wrapper
            // from our collection.
            openRegions.Remove(sender as FormRegionControls);
        }

        public object GetFormRegionIcon(
            string FormRegionName, int LCID, Outlook.OlFormRegionIcon Icon)
        {
            object icon = null;
            switch (Icon)
            {
                // This is a 'separate' form region, so only the page icon is used.
                case Outlook.OlFormRegionIcon.olFormRegionIconPage:
                    icon = PictureConverter.IconToPictureDisp(Resources.page);
                    break;
            }
            return icon;
        }

        public object GetFormRegionManifest(string FormRegionName, int LCID)
        {
            switch (FormRegionName)
            {
                case Connect.FormRegionValue1:
                    return Resources.CustomFormRegion_1_XML;
                case Connect.FormRegionValue2:
                    return Resources.CustomFormRegion_2_XML;
                case Connect.FormRegionValue3:
                    return Resources.CustomFormRegion_3_XML;
                default:
                    return null;
            }
        }

        #endregion

    }

}

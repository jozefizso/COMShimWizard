using System;
using Extensibility;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
//using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Configuration;
using Microsoft.Win32;
using System.Drawing;

namespace TestExcel2007Addin
{

    [Guid(Connect.TypeGuid)]
    [ProgId(Connect.TypeProgId)]
    [ComVisible(true)]
    public class Connect :
        Office.IRibbonExtensibility,
        Extensibility.IDTExtensibility2,
        Office.ICustomTaskPaneConsumer
    {

        #region fields

        private Office.CustomTaskPane taskPane;
        public const string TypeGuid = "42CE11F6-7B06-4836-BBA7-1A5FB233F3C4";
        public const string TypeProgId = "TestExcel2007Addin.Connect";
        internal const string OfficeAddinKeyName =
            @"Software\Microsoft\Office\Excel\Addins\" + Connect.TypeProgId;

        #endregion


        #region Registration

        // We only need to explicitly register these subkeys:

        // [HKCU\Software\Microsoft\Office\Excel\Addins\TestExcel2007Addin.Connect]
        // "FriendlyName"="TestExcel2007Addin"
        // "Description"=""
        // "LoadBehavior"=dword:00000003
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
        }

        [ComUnregisterFunction]
        public static void UnRegisterFunction(Type type)
        {
            Registry.CurrentUser.DeleteSubKey(Connect.OfficeAddinKeyName);
        }

        #endregion


        #region IDTExtensibility2

        public Connect()
        {
        }

        public void OnConnection(
            object application, Extensibility.ext_ConnectMode connectMode,
            object addInInst, ref System.Array custom)
        {
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


        #region IRibbonExtensibility

        public string GetCustomUI(string ribbonID)
        {
            return Resources.RibbonX;
        }

        public void OnTaskPaneToggle(Office.IRibbonControl control, bool isPressed)
        {
            taskPane.Visible = isPressed;
        }

        #endregion


        #region ICustomTaskPaneConsumer

        public void CTPFactoryAvailable(Office.ICTPFactory CTPFactoryInst)
        {
            try
            {
                String taskPaneTitle = ConfigurationManager.AppSettings["taskPaneTitle"];
                if (taskPaneTitle == null)
                {
                    taskPaneTitle = "default";
                }

                taskPane = CTPFactoryInst.CreateCTP(
                    "TestExcel2007Addin.SimpleControl", taskPaneTitle, Type.Missing);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        #endregion

    }

}

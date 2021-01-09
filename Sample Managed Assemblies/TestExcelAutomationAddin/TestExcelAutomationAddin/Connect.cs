using System;
using Extensibility;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;

// Note: must set the build properties to include "Register for COM Interop"
// so that the DLL is COM-registered, and "Make COM Visible" so that we get a
// typelib registered.

namespace TestExcelAutomationAddin
{

    #region Custom interfaces

    [Guid("326D7F7F-7279-4834-BF56-7C98B4B09E05")]
    public interface ITemperatureConversion
    {
        double Fahr2Cel(double val, [Optional] object isVolatile);
        double Cel2Fahr(double val);
    }

    [Guid("7CFB309B-D958-4d73-8CFC-2E8123787ED7")]
    public interface IRangeCalcs
    {
        double SumRangeValues(Excel.Range range);
    }

    [Guid("08ACFF42-FDC5-4838-9FDD-47B043719A32")]
    public interface IAppDomainReport
    {
        string GetAppDomain();
    }

    #endregion

    [Guid(UdfClass.TypeGuid)]
    [ProgId(UdfClass.TypeProgId)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class UdfClass :
        Extensibility.IDTExtensibility2,
        ITemperatureConversion,
        IRangeCalcs,
        IAppDomainReport
    {

        #region Fields

        internal const string TypeGuid = "83430446-60AC-403C-9079-270BD0DFCD8D";
        public const string TypeProgId = "TestExcelAutomationAddin.UdfClass";
        const string SubKeyName = @"CLSID\{" + TypeGuid + @"}\Programmable";

        private Excel.Application xl;
        private const double constantFactor = 32.0;

        #endregion


        #region IDTExtensibility2

        public void OnConnection(
            object application, Extensibility.ext_ConnectMode connectMode,
            object addInInst, ref System.Array custom)
        {
            xl = (Excel.Application)application;
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


        #region Object overrides

        // We override these members because we don't want
        // them showing up in Excel's Insert Function dialog.

        [ComVisible(false)]
        public override bool Equals(object obj)
        {
            return base.Equals(obj);
        }

        [ComVisible(false)]
        public override int GetHashCode()
        {
            return base.GetHashCode();
        }

        [ComVisible(false)]
        public override string ToString()
        {
            return base.ToString();
        }

        // Note: you can't override GetType.

        public UdfClass()
        {
        }

        #endregion


        #region Registration

        // We only need to explicitly register this subkey:
        // HKCR\CLSID\{83430446-60AC-403C-9079-270BD0DFCD8D}\Programmable
        // All other COM keys are created because we specify register for
        // COM interop in the project properties.

        [ComRegisterFunction]
        public static void RegisterFunction(Type type)
        {
            Registry.ClassesRoot.CreateSubKey(UdfClass.SubKeyName);
        }

        [ComUnregisterFunction]
        public static void UnRegisterFunction(Type type)
        {
            Registry.ClassesRoot.DeleteSubKey(UdfClass.SubKeyName);
        }

        #endregion


        #region ITemperatureConversion

        // This function is volatile.
        // The function's calculations depend on the value of cell A1
        // in the worksheet, even though this cell is not referenced
        // either directly or indirectly in the function's argument 
        // list. However, marking the function as volatile ensures 
        // that the function is recalculated whenever any cell changes
        // (including cell A1).
        // The optional parameter has to be an object not a bool,
        // so that we can test its type against System.Reflection.Missing.
        public double Fahr2Cel(double val, [Optional] object isVolatile)
        {
            if (!(isVolatile is System.Reflection.Missing))
            {
                object missing = Type.Missing;
                if ((bool)isVolatile)
                {
                    xl.Volatile(missing);
                }
            }

            // Get the conversion factor from the sheet.
            Excel.Worksheet sheet = (Excel.Worksheet)xl.ActiveSheet;
            double variableFactor =
                (double)((Excel.Range)sheet.Cells[1, 1]).Value2;
            return ((5.0 / 9.0) * (val - variableFactor));
        }

        public double Cel2Fahr(double val)
        {
            return ((val * (9.0 / 5.0)) + constantFactor);
        }

        #endregion


        #region IRangeCalcs

        // Read values from the OM (the given range).
        public double SumRangeValues(Excel.Range range)
        {
            double retVal = 0;
            for (int row = 1; row <= range.Rows.Count; row++)
            {
                for (int col = 1; col <= range.Columns.Count; col++)
                {
                    retVal += Convert.ToDouble(
                        ((Excel.Range)range.Cells[row, col]).Value2);
                }
            }
            return retVal;
        }

        #endregion


        #region IAppDomainReport Members

        public string GetAppDomain()
        {
            return AppDomain.CurrentDomain.FriendlyName;
        }

        #endregion
    }
}

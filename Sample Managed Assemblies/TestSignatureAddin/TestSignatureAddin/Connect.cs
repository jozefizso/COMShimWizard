using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Extensibility;
using Office = Microsoft.Office.Core;
using System.Drawing;
using Microsoft.Win32;

// This is a minimal custom signature provider. To test, this, run Word 2007,
// go to the Insert tab, drop down the Signature Line list.
// The custom entry "TestSignatureAddin" should be in the list.
// This is the add-in's registered friendly name.
namespace TestSignatureAddin
{

    [Guid(Connect.TypeGuid)]
    [ProgId(Connect.TypeProgId)]
    [ComVisible(true)]
    public class Connect : IDTExtensibility2, Office.SignatureProvider
    {
    
        #region IDTExtensibility2

        public Connect()
        {
        }

        public void OnConnection(object application, ext_ConnectMode connectMode,
            object addInInst, ref System.Array custom)
        {
        }
    
        public void OnDisconnection(ext_DisconnectMode disconnectMode, ref System.Array custom)
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

        #endregion IDTExtensibility2

        #region fields

        public const string TypeGuid = "6DF7DA7E-6660-404A-99B0-9B488A06CB24";
        public const string TypeProgId = "TestSignatureAddin.Connect";
        internal const string OfficeAddinKeyName =
            @"Software\Microsoft\Office\Word\Addins\" + Connect.TypeProgId;

        #endregion


        #region Registration

        // We need to explicitly register these Office regkeys:
        //
        // [HKCU\Software\Microsoft\Office\Word\Addins\TestWord2007Addin.Connect]
        // "FriendlyName"="TestWord2007Addin"
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


        #region SignatureProvider

        public object GetProviderDetail(
            Office.SignatureProviderDetail sigprovdet)
        {
            object detail = null;
            switch (sigprovdet)
            {
                case Office.SignatureProviderDetail.sigprovdetHashAlgorithm:
                    detail = "SHA-1";
                    break;
                case Office.SignatureProviderDetail.sigprovdetUIOnly:
                    detail = false;
                    break;
                case Office.SignatureProviderDetail.sigprovdetUrl:
                    detail = "http://www.microsoft.com";
                    break;
                default:
                    detail = "\0";
                    break;
            }
            return detail;
        }

        public void ShowSignatureSetup(
            object ParentWindow, Office.SignatureSetup psigsetup)
        {
            psigsetup.SuggestedSigner = "Andrew";
        }

        public stdole.IPictureDisp GenerateSignatureLineImage(
            Office.SignatureLineImage siglnimg,
            Office.SignatureSetup psigsetup,
            Office.SignatureInfo psiginfo,
            object XmlDsigStream)
        {
            stdole.IPictureDisp picture = null;
            Bitmap b = new Bitmap(200, 50);

            if (siglnimg == Office.SignatureLineImage.siglnimgUnsigned)
            {
                Graphics g = Graphics.FromImage(b);
                g.DrawRectangle(new Pen(Color.Red, 2), 0, 0, 200, 50);
                g.FillRectangle(new SolidBrush(Color.Thistle), 2, 2, 196, 46);
                g.DrawString(String.Format("{0} ({1})",
                    psigsetup.SuggestedSigner, DateTime.Now.ToShortDateString()),
                    new Font("Courier", 12),
                    new SolidBrush(Color.MidnightBlue), new PointF(30, 16));
            }

            picture = PictureConverter.ImageToPictureDisp(
                Image.FromHbitmap(b.GetHbitmap()));
            return picture;
        }

        #region _notimpl SignatureProvider methods

        public Array HashStream(
            object QueryContinue, object Stream)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public void NotifySignatureAdded(
            object ParentWindow, Office.SignatureSetup psigsetup, 
            Office.SignatureInfo psiginfo)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public void ShowSignatureDetails(
            object ParentWindow, Office.SignatureSetup psigsetup, 
            Office.SignatureInfo psiginfo, object XmlDsigStream, 
            ref Office.ContentVerificationResults pcontverres, 
            ref Office.CertificateVerificationResults pcertverres)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public void ShowSigningCeremony(
            object ParentWindow, Office.SignatureSetup psigsetup, 
            Office.SignatureInfo psiginfo)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public void SignXmlDsig(
            object QueryContinue, Office.SignatureSetup psigsetup, 
            Office.SignatureInfo psiginfo, object XmlDsigStream)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public void VerifyXmlDsig(
            object QueryContinue, 
            Office.SignatureSetup psigsetup, 
            Office.SignatureInfo psiginfo, 
            object XmlDsigStream, 
            ref Office.ContentVerificationResults pcontverres, 
            ref Office.CertificateVerificationResults pcertverres)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        #endregion

        #endregion

    }

}
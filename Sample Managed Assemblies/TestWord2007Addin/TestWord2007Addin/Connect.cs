using System;
using Extensibility;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using System.Drawing;
using System.Collections;
using System.Collections.Generic;
using Microsoft.Win32;
using System.Configuration;

namespace TestWord2007Addin
{
    [Guid(Connect.TypeGuid)]
    [ProgId(Connect.TypeProgId)]
    [ComVisible(true)]
    public class Connect :
        Extensibility.IDTExtensibility2,
        Office.ICustomTaskPaneConsumer,
        Office.IRibbonExtensibility,
        Office.IBlogExtensibility,
        Office.IBlogPictureExtensibility,
        Office.SignatureProvider,
        Office.EncryptionProvider,
        Office.IDocumentInspector
    {

        #region fields

        public const string TypeGuid = "F690A224-A3FE-46FF-B38C-2D0DD21D42C4";
        public const string TypeProgId = "TestWord2007Addin.Connect";
        internal const string OfficeAddinKeyName =
            @"Software\Microsoft\Office\Word\Addins\" + Connect.TypeProgId;
        internal const string BlogAddinsKeyName =
            @"Software\Microsoft\Office\Common\Blog\Addins\";
        internal const string EncryptionKeyName =
            @"Software\Microsoft\Office\12.0\Common\Security";
        internal const string EncryptionKeyValue = "OpenXMLEncryption";
        internal const string DocumentInspectorsKeyName =
            @"Software\Microsoft\Office\12.0\Word\Document Inspectors\UK English to US English";

        private Office.CustomTaskPane taskPane;

        private string[] ukWords = new string[] { "colour", "centre", "lorry" };
        private string[] usWords = new string[] { "color", "center", "truck" };

        private string blogProvider;
        private string friendlyName;

        #endregion


        #region Registration

        // We need to explicitly register these Office regkeys:
        //
        // [HKCU\Software\Microsoft\Office\Word\Addins\TestWord2007Addin.Connect]
        // "FriendlyName"="TestWord2007Addin"
        // "Description"=""
        // "LoadBehavior"=dword:00000003
        //
        // [HKCU\Software\Microsoft\Office\Common\Blog]
        // [HKCU\Software\Microsoft\Office\Common\Blog\Addins]
        // "TestWord2007Addin.Connect"=""
        //
        // [HKCU\Software\Microsoft\Office\12.0\Common\Security]
        // "OpenXMLEncryption"="F690A224-A3FE-46FF-B38C-2D0DD21D42C4"
        //
        // [HKLM\Software\Microsoft\Office\12.0\Word\Document Inspectors]
        // [HKLM\Software\Microsoft\Office\12.0\Word\Document Inspectors\UK English to US English]
        // "CLSID"="{F690A224-A3FE-46FF-B38C-2D0DD21D42C4}"
        // "Selected"=dword:00000001
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

            using (RegistryKey blogKey =
                Registry.CurrentUser.CreateSubKey(Connect.BlogAddinsKeyName))
            {
                blogKey.SetValue(Connect.TypeProgId, "");
            }

            using (RegistryKey encryptionKey =
                Registry.CurrentUser.CreateSubKey(Connect.EncryptionKeyName))
            {
                encryptionKey.SetValue(Connect.EncryptionKeyValue, Connect.TypeGuid);
            }

            // Note this key is in HKLM.
            using (RegistryKey documentInspectorsKey =
                Registry.LocalMachine.CreateSubKey(Connect.DocumentInspectorsKeyName))
            {
                documentInspectorsKey.SetValue("CLSID", "{" + Connect.TypeGuid + "}");
                documentInspectorsKey.SetValue("Selected", 1);
            }
        }

        [ComUnregisterFunction]
        public static void UnRegisterFunction(Type type)
        {
            Registry.CurrentUser.DeleteSubKey(Connect.OfficeAddinKeyName);

            using (RegistryKey blogKey =
                Registry.CurrentUser.CreateSubKey(Connect.BlogAddinsKeyName))
            {
                blogKey.DeleteValue(Connect.TypeProgId);
            }

            using (RegistryKey encryptionKey =
                Registry.CurrentUser.CreateSubKey(Connect.EncryptionKeyName))
            {
                encryptionKey.DeleteValue(Connect.EncryptionKeyValue);
            }

            // Note this key is in HKLM.
            Registry.LocalMachine.DeleteSubKey(Connect.DocumentInspectorsKeyName);
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

        public void OnTaskPaneToggle(
            Office.IRibbonControl control, bool isPressed)
        {
            taskPane.Visible = isPressed;
        }

        #endregion


        #region ICustomTaskPaneConsumer

        public void CTPFactoryAvailable(Office.ICTPFactory CTPFactoryInst)
        {
            try
            {
                String taskPaneTitle = 
                    ConfigurationManager.AppSettings["taskPaneTitle"];
                if (taskPaneTitle == null)
                {
                    taskPaneTitle = "default";
                }

                taskPane = CTPFactoryInst.CreateCTP(
                    "TestWord2007Addin.SimpleControl", 
                    taskPaneTitle, Type.Missing);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        #endregion


        #region EncryptionProvider

        public object GetProviderDetail(
            Office.EncryptionProviderDetail encprovdet)
        {
            object detail = null;
            switch (encprovdet)
            {
                case Office.EncryptionProviderDetail.encprovdetUrl:
                    detail = "http://www.microsoft.com";
                    break;
                case Office.EncryptionProviderDetail.encprovdetAlgorithm:
                    detail = "XOR";
                    break;
                case Office.EncryptionProviderDetail.encprovdetBlockCipher:
                    detail = false;
                    break;
                case Office.EncryptionProviderDetail.encprovdetCipherBlockSize:
                case Office.EncryptionProviderDetail.encprovdetCipherMode:
                    detail = 0;
                    break;
            }
            return detail;
        }

        public int NewSession(object ParentWindow)
        {
            MessageBox.Show("TestWord2007Addin.Connect");
            return 0;
        }


        #region _notimpl methods

        public int Authenticate(object ParentWindow, object EncryptionData, out uint PermissionsMask)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public int CloneSession(int SessionHandle)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public void DecryptStream(int SessionHandle, string StreamName, object EncryptedStream, object UnencryptedStream)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public void EncryptStream(int SessionHandle, string StreamName, object UnencryptedStream, object EncryptedStream)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public void EndSession(int SessionHandle)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public int Save(int SessionHandle, object EncryptionData)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public void ShowSettings(int SessionHandle, object ParentWindow, bool ReadOnly, out bool Remove)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        #endregion

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
                    psigsetup.SuggestedSigner, 
                    DateTime.Now.ToShortDateString()),
                    new Font("Courier", 12),
                    new SolidBrush(Color.MidnightBlue), 
                    new PointF(30, 16));
            }

            picture = PictureConverter.ImageToPictureDisp(
                Image.FromHbitmap(b.GetHbitmap()));
            return picture;
        }

        #region _notimpl methods

        public Array HashStream(
            object QueryContinue, object Stream)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public void NotifySignatureAdded(
            object ParentWindow, Office.SignatureSetup psigsetup, Office.SignatureInfo psiginfo)
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
            object ParentWindow, Office.SignatureSetup psigsetup, Office.SignatureInfo psiginfo)
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
            object QueryContinue, Office.SignatureSetup psigsetup, Office.SignatureInfo psiginfo,
            object XmlDsigStream, ref Office.ContentVerificationResults pcontverres,
            ref Office.CertificateVerificationResults pcertverres)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        #endregion

        #endregion


        #region IBlogExtensibility

        public void BlogProviderProperties(
            out string BlogProvider, out string FriendlyName,
            out Microsoft.Office.Core.MsoBlogCategorySupport CategorySupport,
            out bool Padding)
        {
            BlogProvider = blogProvider = "Contoso Blog Provider";
            FriendlyName = friendlyName = "Contoso Blog FriendlyName";
            CategorySupport = 
                Office.MsoBlogCategorySupport.msoBlogNoCategories;
            Padding = false;
        }

        public void SetupBlogAccount(string Account, int ParentWindow,
            object Document, bool NewAccount, out bool ShowPictureUI)
        {
            ShowPictureUI = true;

            RegistryKey regAccount = Registry.CurrentUser.OpenSubKey(
                "Software\\Microsoft\\Office\\Common\\Blog\\Accounts\\"
                + Account, true);
            regAccount.SetValue(
                "FriendlyName", friendlyName, RegistryValueKind.String);
            regAccount.SetValue(
                "Provider", blogProvider, RegistryValueKind.String);
        }

        public void GetUserBlogs(string Account, int ParentWindow,
            object Document, out Array BlogNames, out Array BlogIDs,
            out Array BlogURLs)
        {
            BlogNames = new String[] { "BlogName1" };
            BlogIDs = new String[] { "1234567890" };
            BlogURLs = new String[] { "www.contoso.com" };
        }

        #region _notimpl methods

        public void GetCategories(string Account, int ParentWindow,
            object Document, out Array Categories)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public void GetRecentPosts(string Account, int ParentWindow,
            object Document, out Array PostTitles, out Array PostDates,
            out Array PostIDs)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public void Open(string Account, string PostID, int ParentWindow,
            out string xHTML, out string Title, out string DatePosted,
            out Array Categories)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public void PublishPost(string Account, int ParentWindow,
            object Document, string xHTML, string Title, string DateTime,
            Array Categories, bool Draft, out string PostID,
            out string PublishMessage)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public void RepublishPost(string Account, int ParentWindow,
            object Document, string PostID, string xHTML, string Title,
            string DateTime, Array Categories, bool Draft,
            out string PublishMessage)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        #endregion

        #endregion


        #region IBlogPictureExtensibility

        public void BlogPictureProviderProperties(
            out string BlogPictureProvider, out string FriendlyName)
        {
            BlogPictureProvider = blogProvider;
            FriendlyName = friendlyName;
        }

        public void CreatePictureAccount(string Account, string BlogProvider,
            int ParentWindow, object Document)
        {
            RegistryKey regAccount = Registry.CurrentUser.OpenSubKey(
                "Software\\Microsoft\\Office\\Common\\Blog\\Accounts\\"
                + Account, true);

            regAccount.SetValue(
                "ImagePublishURL", "www.contoso.com", RegistryValueKind.String);
            regAccount.SetValue(
                "ImagePublishedURL", "www.contoso.com", RegistryValueKind.String);
        }

        #region _notimpl method

        public void PublishPicture(string Account, int ParentWindow,
            object Document, object Image, out string PictureURI,
            int ImageType)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        #endregion

        #endregion


        #region IDocumentInspector

        // These strings are displayed in the Inspect Document Dialog.
        // The Name is the checkbox label.
        // The Desc is the description below.
        public void GetInfo(out string Name, out string Desc)
        {
            Name = "UK to US Converter";
            Desc = "Replaces UK spelling with US spelling.";
        }

        // Invoked when the user chooses to go ahead and inspect the document.
        public void Inspect(object Doc, out Office.MsoDocInspectorStatus Status,
            out string Result, out string Action)
        {
            Word.Document document = null;
            Word.StoryRanges storyRanges = null;
            Word.Range range = null;

            try
            {
                document = (Word.Document)Doc;
                storyRanges = document.StoryRanges;

                // Build a list of words in the document that match the 
                // known UK words in our list.
                ArrayList itemsFound = new ArrayList();
                object find = null;
                object match = true;
                object missing = Type.Missing;
                foreach (string ukWord in this.ukWords)
                {
                    find = ukWord;
                    // We need to keep resetting the range to the whole document,
                    // because Execute will reset it.
                    range = storyRanges[Word.WdStoryType.wdMainTextStory];
                    if (range.Find.Execute(
                        ref find, ref match,
                        ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing,
                        ref missing))
                    {
                        itemsFound.Add(ukWord);
                    }
                }

                if (itemsFound.Count > 0)
                {
                    Status = 
                        Office.MsoDocInspectorStatus.msoDocInspectorStatusIssueFound;
                    Result =
                        String.Format("{0} UK words found.", itemsFound.Count);
                    Action = "Replace UK Words";
                }
                else
                {
                    Status = Office.MsoDocInspectorStatus.msoDocInspectorStatusDocOk;
                    Result = "No UK words found";
                    Action = "No UK words to remove";
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.ToString());
                Status = Office.MsoDocInspectorStatus.msoDocInspectorStatusError;
                Result = "Error.";
                Action = "No action.";
            }
            finally
            {
                // This is important: if we don't clean up explicitly,
                // the Word process will not terminate.
                storyRanges = null;
                range = null;
                document = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        // Invoked when the user chooses to fix the items found by the
        // Inspect operation.
        public void Fix(object Doc, int Hwnd,
            out Office.MsoDocInspectorStatus Status, out string Result)
        {
            Word.Document document = null;
            Word.StoryRanges storyRanges = null;
            Word.Range range = null;

            try
            {
                document = (Word.Document)Doc;
                storyRanges = document.StoryRanges;
                range = storyRanges[Word.WdStoryType.wdMainTextStory];

                object find = null;
                object matchCase = false;
                object matchWord = true;
                object replaceWith = null;
                object matchWholeWord = false;
                object replace = Word.WdReplace.wdReplaceAll;
                object missing = Type.Missing;
                int i = 0;

                // Scan the whole document text and execute a search+replace
                // for all words that match our list of known UK spellings.
                foreach (string ukWord in ukWords)
                {
                    find = ukWord;
                    replaceWith = usWords[i];
                    range.Find.Execute(
                        ref find, ref matchCase, ref matchWholeWord, ref missing,
                        ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref replaceWith, ref replace, ref missing,
                        ref missing, ref missing, ref missing);
                    i++;
                }
                Status = Office.MsoDocInspectorStatus.msoDocInspectorStatusDocOk;
                Result = "All UK words have been replaced.";
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.ToString());
                Status = Office.MsoDocInspectorStatus.msoDocInspectorStatusError;
                Result = "Error.";
            }
            finally
            {
                storyRanges = null;
                range = null;
                document = null;

                // Note that if we force a GC collect here, it will hang Word.
                //GC.Collect();
                //GC.WaitForPendingFinalizers();
                //GC.Collect();
                //GC.WaitForPendingFinalizers();
            }
        }

        #endregion

    }

}

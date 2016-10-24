using System;
using Extensibility;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using Microsoft.Win32;

// This is a very minimal (incomplete) custom blog add-in. It has just enough
// functionality to confirm that the add-in gets loaded and called through 
// the IBlogExtensibility and IBlogPictureExtensibility interfaces.

// To test this minimal functionality, run Word, select File | New | Blog Post.
// Select Manage Accounts. In the Blog Accounts dialog, you will see a row
// for the Simple Blog FriendlyName.

namespace TestBlogAddin
{
    [Guid(Connect.TypeGuid)]
    [ProgId(Connect.TypeProgId)]
    [ComVisible(true)]
	public class Connect : 
        Object, 
        Extensibility.IDTExtensibility2,
        Office.IBlogExtensibility, 
        Office.IBlogPictureExtensibility

	{
		public Connect()
		{
        }

        #region Fields

        public const string TypeGuid = "07353DE8-D782-4750-B04E-A037D12304CD";
        public const string TypeProgId = "TestBlogAddin.Connect";
        internal const string OfficeAddinKeyName =
            @"Software\Microsoft\Office\Word\Addins\" + Connect.TypeProgId;
        internal const string BlogAddinsKeyName =
            @"Software\Microsoft\Office\Common\Blog\Addins\";
        private string blogProvider;
        private string friendlyName;

        #endregion


        #region Registration

        // We need to explicitly register these Office regkeys:
        //
        // [HKCU\Software\Microsoft\Office\Word\Addins\TestBlogAddin.Connect]
        // "FriendlyName"="TestBlogAddin"
        // "Description"=""
        // "LoadBehavior"=dword:00000003
        //
        // [HKCU\Software\Microsoft\Office\Common\Blog]
        // [HKCU\Software\Microsoft\Office\Common\Blog\Addins]
        // "TestBlogAddin.Connect"=""
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
        }

        #endregion


        #region IDTExtensibility2

        public void OnConnection(object application,
            Extensibility.ext_ConnectMode connectMode,
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


        #region IBlogExtensibility Members

        public void BlogProviderProperties(
            out string BlogProvider, out string FriendlyName, 
            out Microsoft.Office.Core.MsoBlogCategorySupport CategorySupport, 
            out bool Padding)
        {
            BlogProvider = blogProvider = "Simple Blog Provider";
            FriendlyName = friendlyName = "Simple Blog FriendlyName";
            CategorySupport = Office.MsoBlogCategorySupport.msoBlogNoCategories;
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

        #region _notimpl Members

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


        #region IBlogPictureExtensibility Members

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

            regAccount.SetValue("ImagePublishURL", "www.contoso.com", RegistryValueKind.String);
            regAccount.SetValue("ImagePublishedURL", "www.contoso.com", RegistryValueKind.String);
        }

        #region _notimpl Member

        public void PublishPicture(string Account, int ParentWindow, 
            object Document, object Image, out string PictureURI, 
            int ImageType)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        #endregion

        #endregion
    }
}
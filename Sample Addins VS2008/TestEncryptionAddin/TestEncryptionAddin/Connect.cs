using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Extensibility;
using Office = Microsoft.Office.Core;
using Microsoft.Win32;

// This is a minimal custom encryption provider. To test, this, run Word 2007,
// go to the File menu, select Prepare | Encrypt Document.
// The result will be the messagebox from this add-in's NewSession method.
namespace TestEncryptionAddin
{
    [Guid(Connect.TypeGuid)]
    [ProgId(Connect.TypeProgId)]
    [ComVisible(true)]
    public class Connect : IDTExtensibility2, Office.EncryptionProvider
    {

        #region fields

        public const string TypeGuid = "1F6E6CC6-FC04-4F60-A673-7DFD1715C1C4";
        public const string TypeProgId = "TestEncryptionAddin.Connect";
        internal const string OfficeAddinKeyName =
            @"Software\Microsoft\Office\Word\Addins\" + Connect.TypeProgId;
        internal const string EncryptionKeyName =
            @"Software\Microsoft\Office\12.0\Common\Security";
        internal const string EncryptionKeyValue = "OpenXMLEncryption";

        #endregion


        #region Registration

        // We need to explicitly register these Office regkeys:
        //
        // [HKCU\Software\Microsoft\Office\Word\Addins\TestEncryptionAddin.Connect]
        // "FriendlyName"="TestEncryptionAddin"
        // "Description"=""
        // "LoadBehavior"=dword:00000003
        //
        // [HKCU\Software\Microsoft\Office\12.0\Common\Security]
        // "TestEncryptionAddin.Connect"="1F6E6CC6-FC04-4F60-A673-7DFD1715C1C4"
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

            using (RegistryKey encryptionKey =
                Registry.CurrentUser.CreateSubKey(Connect.EncryptionKeyName))
            {
                encryptionKey.SetValue(Connect.EncryptionKeyValue, Connect.TypeGuid);
            }
        }

        [ComUnregisterFunction]
        public static void UnRegisterFunction(Type type)
        {
            Registry.CurrentUser.DeleteSubKey(Connect.OfficeAddinKeyName);

            using (RegistryKey encryptionKey =
                Registry.CurrentUser.CreateSubKey(Connect.EncryptionKeyName))
            {
                encryptionKey.DeleteValue(Connect.EncryptionKeyValue);
            }
        }

        #endregion


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


        #region EncryptionProvider

	    // This is where a provider provides metadata about itself.
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

        // Office invokes the NewSession method when the user selects
        // Prepare | Encrypt Document.
        // This is where a provider shows any UI appropriate for applying 
        // encryption. For example, a password encryptor would prompt for 
        // the user's password. In this sample implementation, we use an
        // arbitrary string used for trivial XOR obfuscation.
        public int NewSession(object ParentWindow)
        {
            MessageBox.Show("TestEncryptionAddin.Connect.NewSession");
            return 0;
        }

        #region _notimpl methods

        // After calling NewSession, if the user returns to Prepare | Encrypt,
        // Office calls ShowSettings. We can show whatever UI we like here,
        // and either set the document to be encrypted, or decrypted.
        public void ShowSettings(int SessionHandle, object ParentWindow, 
            bool ReadOnly, out bool Remove)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        // When the document is saved (by the user or autosave), Office 
        // performs these steps:
        // 1. Calls CloneSession to create a second, working copy of our
        //    encryption session for the file that is about to be saved.
        // 2. Calls the Save method to get whatever custom information 
        //    we want to persist about the encryption settings. This 
        //    information will be returned to us when this document is 
        //    reopened later.
        // 3. Calls EncryptStream and hands our provider the entire contents
        //    of the document. We may apply whatever encryption we like.
        // 4. Calls EndSession on the cloned session handle.

        // This is where a provider makes a copy of the session data (eg used 
        // when performing a background save or when creating a new file based
        // on an existing encrypted file).
        // Note: this runs on a background thread, so we must not pop modal
        // UI during this call, or we'll crash Office.
        public int CloneSession(int SessionHandle)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        // This is where a provider stores information it needs to maintain a 
        // session. This data gets serialized as a base64-encoded blob inside
        // the EncryptionInfo stream. For example, a password encryptor would
        // store the password verifier. 
        public int Save(int SessionHandle, object EncryptionData)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        // This is where a provider encrypts the data in a stream. For example,
        // a password encryptor would generate a key based on the password and
        // use an algorithm such as AES128 to encrypt the data. 
        // Note: this runs on a background thread, so we must not pop modal
        // UI during this call, or we'll crash Office.
        public void EncryptStream(int SessionHandle, string StreamName, 
            object UnencryptedStream, object EncryptedStream)
        {
            throw new Exception("The method or operation is not implemented.");
        }

	    // This is where a provider cleans up resources for a session.
        public void EndSession(int SessionHandle)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        // This is where a provider shows suitable UI for determining whether
        // the user has access to the document and/or specifying what rights
        // the user has. For example, a password encryptor would prompt for
        // the document's password and generate a decryption key.
        public int Authenticate(object ParentWindow, object EncryptionData, 
            out uint PermissionsMask)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        // This is where a provider decrypts the data in a stream. For example,
        // a password encryptor would generate a key based on the password and
        // use an algorithm such as AES128 to decrypt the data.
        public void DecryptStream(int SessionHandle, string StreamName, 
            object EncryptedStream, object UnencryptedStream)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        #endregion

        #endregion

    }
}
using System;
using Extensibility;
using System.Runtime.InteropServices;
using System.Collections;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Win32;

namespace TestDocumentInspector
{

    [Guid(Connect.TypeGuid)]
    [ProgId(Connect.TypeProgId)]
    [ComVisible(true)]
    public class Connect : 
        Object, Extensibility.IDTExtensibility2,
        Office.IDocumentInspector
	{

        #region fields

        public const string TypeGuid = "634E508D-EF84-407A-ABBE-767B82AE7798";
        public const string TypeProgId = "TestDocumentInspector.Connect";
        internal const string OfficeAddinKeyName =
            @"Software\Microsoft\Office\Word\Addins\" + Connect.TypeProgId;
        internal const string DocumentInspectorsKeyName =
            @"Software\Microsoft\Office\12.0\Word\Document Inspectors\UK English to US English";

        private string[] ukWords = new string[] { "colour", "centre", "lorry" };
        private string[] usWords = new string[] { "color", "center", "truck" };

        #endregion


        #region Registration

        // We need to explicitly register these Office regkeys:
        //
        // [HKCU\Software\Microsoft\Office\Word\Addins\TestDocumentInspector.Connect]
        // "FriendlyName"="TestDocumentInspector"
        // "Description"=""
        // "LoadBehavior"=dword:00000003
        //
        // [HKLM\Software\Microsoft\Office\12.0\Word\Document Inspectors]
        // [HKLM\Software\Microsoft\Office\12.0\Word\Document Inspectors\UK English to US English]
        // "CLSID"="{634E508D-EF84-407A-ABBE-767B82AE7798}"
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


        #region IDocumentInspector Members

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
                object match = false;
                object missing = Type.Missing;
                foreach (string ukWord in this.ukWords)
                {
                    find = ukWord;
                    // We need to keep resetting the range to the whole document,
                    // because Execute will reset it.
                    range = storyRanges[Word.WdStoryType.wdMainTextStory];
                    if (range.Find.Execute(
                        ref find, ref match, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing))
                    {
                        itemsFound.Add(ukWord);
                    }
                }

                if (itemsFound.Count > 0)
                {
                    Status = Office.MsoDocInspectorStatus.msoDocInspectorStatusIssueFound;
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
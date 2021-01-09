using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Collections;

using SmartTagLib = Microsoft.Office.Interop.SmartTag;


namespace TestSmartTagRecognizer
{

	[GuidAttribute("EDB046FC-9025-42cd-BFF8-A0D7E629C478")]
	[ProgId("TestSmartTagRecognizer.SomeProgId")]
	public class Recognizer : SmartTagLib.ISmartTagRecognizer
	{

		private string[] numericStrings = {"one", "two", "three"};

		// ISmartTagRecognizer ////////////////////////////
		public string ProgId
		{
			get 
			{
				// Return the ProgID of the Recognizer interface.
				return "TestSmartTagRecognizer.Recognizer";
			}
		}

		public int SmartTagCount
		{
			get
			{
				return 1;
			}
		}

		public string get_Desc (int LocaleID)
		{
			return "TestSmartTagRecognizer recognizes company names";
		}

		public string get_Name(int LocaleID)
		{
			return "TestSmartTagRecognizer Recognizer";
		}

		public string get_SmartTagDownloadURL(int SmartTagID)
		{
			return null;
		}

		public string get_SmartTagName(int SmartTagID)
		{
			// This method is called the same number of times as we
			// return in SmartTagCount. This method sets a unique name
			// for the Smart Tag.
			return "Contoso/TestSmartTagRecognizer#Symbols";
		}

		public void Recognize(
			string Text, 
			SmartTagLib.IF_TYPE DataType, 
			int LocaleID, 
			SmartTagLib.ISmartTagRecognizerSite RecognizerSite)
		{
			// The Recognize method is called and passed a text value.
			// We must search for recognized strings in the text in order
			// to set up the set of possible actions.
			int i;
			int startpos;
			int strlen;
			SmartTagLib.ISmartTagProperties propbag;

			int count = numericStrings.Length; 
			for (i = 0; i < count; i++) 
			{
				// See if this Text string matches any groceries.
				startpos = Text.IndexOf(numericStrings[i]) +1;
				strlen = numericStrings[i].Length;
				while (startpos > 0) 
				{
					// Commit the Smart Tag to the property bag.
					propbag = RecognizerSite.GetNewPropertyBag();
					RecognizerSite.CommitSmartTag(
						"Contoso/TestSmartTagRecognizer#Symbols", 
						startpos, strlen, propbag);

					// Continue looking for matches.
					startpos = Text.IndexOf(numericStrings[i], startpos +strlen) +1;
				}
			} 
		}
	
	}
} 






























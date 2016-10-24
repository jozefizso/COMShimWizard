using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;

using SmartTagLib = Microsoft.Office.Interop.SmartTag;


namespace TestSmartTagRecognizerAction
{

	[GuidAttribute("D47D6BAA-3989-46e8-ADCA-250DDE710DF2")]
	[ProgId("TestSmartTagRecognizerAction.SomeActionProgId")]
	public class Action: SmartTagLib.ISmartTagAction
	{

		private string[] numericStrings = {"one", "two", "three"};
		private string[] numericSymbols = {"1", "2", "3"};

		// ISmartTagAction ////////////////////////////////
		public string ProgId
		{
			get 
			{
				return "TestSmartTagRecognizerAction.Action";
			}
		}

		public int SmartTagCount
		{
			get
			{
				return 1;
			}
		}

		public string get_Desc(int LocaleID)
		{
			return "Provides actions for the TestSmartTagRecognizerAction SmartTag";
		}

		public string get_Name(int LocaleID)
		{
			return "TestSmartTagRecognizerAction Smart Tag";
		}

		public string get_SmartTagCaption(int SmartTagID , int LocaleID)
		{
			// This caption is displayed on the menu for the smart tag.
			return "TestSmartTagRecognizerAction Smart Tag";
		}

		public string get_SmartTagName(int SmartTagID)
		{
			// This method is called the same number of times as we
			// return in SmartTagCount. This method sets a unique name
			// for the smart tag.
			return "Contoso/TestSmartTagRecognizerAction#Symbols";
		}

		public string get_VerbCaptionFromID(
			int VerbID, string ApplicationName, int LocaleID)
		{
			// Get a caption for each verb. This caption is displayed
			// on the Smart Tag menu.
			switch(VerbID) 
			{
				case 1:
					return "Symbol for this item";
				default:
					return null;
			}
			//return String.Empty;
		}

		public int get_VerbCount(string SmartTagName)
		{
			// Return the number of verbs we support.
			if (SmartTagName.Equals("Contoso/TestSmartTagRecognizerAction#Symbols")) 
			{
				return 1;
			}
			return 0;
		}

		public int get_VerbID(
			string SmartTagName, int VerbIndex)
		{
			// Return a unique ID for each verb we support.
			return VerbIndex;
		}

		public string get_VerbNameFromID(int VerbID)
		{
			// Return a string name for each verb.
			switch(VerbID) 
			{
				case 1:
					return "Symbol";
				default:
					return null;
			}
		}

		public void InvokeVerb(int VerbID, 
			string ApplicationName, object Target, 
			SmartTagLib.ISmartTagProperties Properties, 
			string Text, string Xml)
		{
			// This method is called when a user invokes a verb
			// from the Smart Tag menu.

			for (int i = 0; i < numericStrings.Length; i++)
			{
				string numericString = numericStrings[i];
				if (String.Compare(numericString, Text, true) == 0) 
				{
					switch(VerbID) 
					{
						case 1:
							// The user wants to show the Description;
							MessageBox.Show("Symbol: " +numericSymbols[i]);
							break;
					}
				}
			}
		}

	}
} 
























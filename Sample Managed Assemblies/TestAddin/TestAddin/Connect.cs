namespace TestAddin
{
	using System;
	using Microsoft.Office.Core;
	using Extensibility;
	using System.Runtime.InteropServices;
	using System.Windows.Forms;
	using System.Diagnostics;
	

	#region Read me for Add-in installation and setup information.
	// When run, the Add-in wizard prepared the registry for the Add-in.
	// At a later time, if the Add-in becomes unavailable for reasons such as:
	//   1) You moved this project to a computer other than which is was originally created on.
	//   2) You chose 'Yes' when presented with a message asking if you wish to remove the Add-in.
	//   3) Registry corruption.
	// you will need to re-register the Add-in by building the MyAddin21Setup project 
	// by right clicking the project in the Solution Explorer, then choosing install.
	#endregion
	

	/// <summary>
	///   The object for implementing an Add-in.
	/// </summary>
	/// <seealso class='IDTExtensibility2' />
	[Guid("7F20788F-7F95-43B4-8224-B1DB54CD7FB7")]
	[ProgId("TestAddin.SomeProgId")]
	public class Connect : Object, Extensibility.IDTExtensibility2
	{

        private object applicationObject;
        private object addInInstance;
        private CommandBarButton button;
        private CommandBar bar;
        private const string buttonName = "TestButton";
        private string hostName = String.Empty;


		/// <summary>
		///		Implements the constructor for the Add-in object.
		///		Place your initialization code within this method.
		/// </summary>
		public Connect()
		{
		}


		/// <summary>
		///      Implements the OnConnection method of the IDTExtensibility2 interface.
		///      Receives notification that the Add-in is being loaded.
		/// </summary>
		/// <param term='application'>
		///      Root object of the host application.
		/// </param>
		/// <param term='connectMode'>
		///      Describes how the Add-in is being loaded.
		/// </param>
		/// <param term='addInInst'>
		///      Object representing this Add-in.
		/// </param>
		/// <seealso class='IDTExtensibility2' />
		public void OnConnection(object application, 
			Extensibility.ext_ConnectMode connectMode, object addInInst, ref System.Array custom)
		{
			applicationObject = application;
			addInInstance = addInInst;

			if (connectMode != Extensibility.ext_ConnectMode.ext_cm_Startup)
			{
				Debug.WriteLine("OnConnection");
				OnStartupComplete(ref custom);
			}
			else
			{
				Debug.WriteLine("OnConnection - ext_cm_Startup");
			}
		}


		/// <summary>
		///     Implements the OnDisconnection method of the IDTExtensibility2 interface.
		///     Receives notification that the Add-in is being unloaded.
		/// </summary>
		/// <param term='disconnectMode'>
		///      Describes how the Add-in is being unloaded.
		/// </param>
		/// <param term='custom'>
		///      Array of parameters that are host application specific.
		/// </param>
		/// <seealso class='IDTExtensibility2' />
		public void OnDisconnection(
            Extensibility.ext_DisconnectMode disconnectMode, ref System.Array custom)
		{
			if (disconnectMode != Extensibility.ext_DisconnectMode.ext_dm_HostShutdown)
			{
				Debug.WriteLine("OnDisconnection");

                if (button != null)
                {
                    try
                    {
                        // We must unhook the event delegate, otherwise we'd end up
                        // with multiple invocations if the user disconnects/reconnects
                        // the add-in.
                        button.Click -=
                            new _CommandBarButtonEvents_ClickEventHandler(button_Click);

                        // We must delete the button when the user disconnects the add-in
                        // during a session. We don't need to do this when the host shuts
                        // down because we create the button as temporary (so it gets
                        // deleted on shutdown anyway).

                        // Deleting the button works fine for most Office apps.
                        // However, there's a race condition in Word. 
                        // Unhooking the event delegate always works, but _any_
                        // subsequent access to the button sometimes throws 0x800A01A8,
                        // regardless of whether or not a shim is used. The workaround 
                        // is to get the button again so that we can delete it.

                        if (hostName == "Microsoft Word")
                        {
                            object missing = Type.Missing;
                            button = (CommandBarButton)
                                bar.FindControl(MsoControlType.msoControlButton,
                                missing, buttonName, true, true);
                        }

                        // We do this for all Office hosts.
                        button.Delete(false);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                }
			}
			else
			{
				Debug.WriteLine("OnDisconnection - ext_dm_HostShutdown");
			}
			applicationObject = null;
		}


		/// <summary>
		///      Implements the OnAddInsUpdate method of the IDTExtensibility2 interface.
		///      Receives notification that the collection of Add-ins has changed.
		/// </summary>
		/// <param term='custom'>
		///      Array of parameters that are host application specific.
		/// </param>
		/// <seealso class='IDTExtensibility2' />
		public void OnAddInsUpdate(ref System.Array custom)
		{
			Debug.WriteLine("OnAddInsUpdate");
		}


        /// <summary>
        ///      Implements the OnBeginShutdown method of the IDTExtensibility2 interface.
        ///      Receives notification that the host application is being unloaded.
        /// </summary>
        /// <param term='custom'>
        ///      Array of parameters that are host application specific.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        public void OnBeginShutdown(ref System.Array custom)
        {
            Debug.WriteLine("OnBeginShutdown");
        }
		

		/// <summary>
		///      Implements the OnStartupComplete method of the IDTExtensibility2 interface.
		///      Receives notification that the host application has completed loading.
		/// </summary>
		/// <param term='custom'>
		///      Array of parameters that are host application specific.
		/// </param>
		/// <seealso class='IDTExtensibility2' />
		public void OnStartupComplete(ref System.Array custom)
		{
			Debug.WriteLine("OnStartupComplete");

			try
			{
				// We want to get the Tools menu, but first we must
                // find the name of the host application, because the
                // way we get to the Tools menu is different depending
                // on which host application we're running in.
				Type applicationType = applicationObject.GetType();
                CommandBars bars = null;
                hostName = (string)applicationType.InvokeMember(
                        "Name", System.Reflection.BindingFlags.GetProperty,
                        null, applicationObject, null);

                if (hostName == "Outlook")
                {
                    // For Outlook, we get to the CommandBars via the ActiveExplorer.
                    object explorer = applicationType.InvokeMember(
                        "ActiveExplorer", System.Reflection.BindingFlags.GetProperty,
                        null, applicationObject, null);
                    if (explorer != null)
                    {
                        Type explorerType = explorer.GetType();
                        bars = (CommandBars)explorerType.InvokeMember(
                            "CommandBars", System.Reflection.BindingFlags.GetProperty,
                            null, explorer, null);
                    }
                }
                else if (hostName == "Microsoft Office InfoPath")
                {
                    // For InfoPath, we get to the CommandBars via the ActiveWindow.
                    object window = applicationType.InvokeMember(
                        "ActiveWindow", System.Reflection.BindingFlags.GetProperty,
                        null, applicationObject, null);
                    if (window != null)
                    {
                        Type windowType = window.GetType();
                        bars = (CommandBars)windowType.InvokeMember(
                            "CommandBars", System.Reflection.BindingFlags.GetProperty,
                            null, window, null);
                    }
                }
                else
                {
                    // For all other Office apps, we get to the CommandBars directly
                    // from the application object.
                    bars = (CommandBars)applicationType.InvokeMember(
                        "CommandBars", System.Reflection.BindingFlags.GetProperty,
                        null, applicationObject, null);
                }
				bar = bars["Tools"];

				// Add our custom button to the bar, if it's not already there.
				object missing = Type.Missing;
                button = (CommandBarButton)
					bar.FindControl(MsoControlType.msoControlButton, 
					missing, buttonName, true, true);
                if (button == null)
				{
					// It's not there, so add our button to the commandbar.
					button = (CommandBarButton) bar.Controls.Add(
						MsoControlType.msoControlButton, 1, missing, missing, true);
					button.FaceId = 59;
					button.Caption = buttonName;
					button.Tag = buttonName;
                    button.Style = MsoButtonStyle.msoButtonIconAndCaption;
				}

				button.Click += 
					new _CommandBarButtonEvents_ClickEventHandler(button_Click);
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.ToString());
			}
		}


		// Dummy handler for the Click event fired when the user clicks our
		// custom CommandBarButton.
		private void button_Click(
			CommandBarButton cmdBarbutton, ref bool cancelButton)
		{
			MessageBox.Show("Hello World");
		}

	}
}




















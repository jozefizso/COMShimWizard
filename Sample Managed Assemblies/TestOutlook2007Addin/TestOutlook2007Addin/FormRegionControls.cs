using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Vbe.Interop.Forms;

namespace TestOutlook2007Addin
{
    // The FormRegionControls class wraps the references to the controls on the
    // custom form region. We'll instantiate a fresh instance of this class for
    // each custom form region that gets opened. This way, we ensure that any
    // UI response is specific to this instance (eg, when the user clicks our
    // commandbutton, we can fetch the textbox text for this same instance.
    public class FormRegionControls
    {
        private Outlook.OlkTextBox textBox1;
        private Outlook.OlkCommandButton commandButton1;
        public event EventHandler Close;

        public FormRegionControls(Outlook.FormRegion region)
        {
            // Fetch the controls from the form, to initialize our managed references.
            UserForm form = region.Form as UserForm;
            Controls formControls = form.Controls;
            textBox1 = formControls.Item("TextBox1") as Outlook.OlkTextBox;
            commandButton1 = formControls.Item("CommandButton1") as Outlook.OlkCommandButton;
            commandButton1.Click += new Outlook.OlkCommandButtonEvents_ClickEventHandler(commandButton1_Click);
            region.Close += new Outlook.FormRegionEvents_CloseEventHandler(region_Close);
        }

        void region_Close()
        {
            // Unhook the Click event sink, clean up all OM references, and notify our parent we're closing.
            commandButton1.Click -= new Outlook.OlkCommandButtonEvents_ClickEventHandler(commandButton1_Click);
            commandButton1 = null;
            textBox1 = null;
            if (Close != null)
            {
                Close(this, EventArgs.Empty);
            }
        }

        void commandButton1_Click()
        {
            MessageBox.Show(textBox1.Text);
        }

    }
}

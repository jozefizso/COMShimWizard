using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace TestOutlook2007Addin
{
    [ComVisible(true)]
    [ProgId("TestOutlook2007Addin.SimpleControl")]
    [Guid("2701C539-77AB-47a1-A65A-97936A8BE6D1")]
    public partial class SimpleControl : UserControl
    {
        public SimpleControl()
        {
            InitializeComponent();
            label.Text += AppDomain.CurrentDomain.FriendlyName;
        }
    }
}

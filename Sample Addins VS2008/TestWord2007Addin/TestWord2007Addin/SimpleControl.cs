using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace TestWord2007Addin
{
    [ComVisible(true)]
    [ProgId("TestWord2007Addin.SimpleControl")]
    [Guid("D8402E38-8E03-4d24-A381-599D526B3E3D")]
    public partial class SimpleControl : UserControl
    {
        public SimpleControl()
        {
            InitializeComponent();
            label.Text += AppDomain.CurrentDomain.FriendlyName;
        }
    }
}

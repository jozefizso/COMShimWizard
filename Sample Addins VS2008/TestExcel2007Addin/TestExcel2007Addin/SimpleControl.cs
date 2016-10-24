using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace TestExcel2007Addin
{
    [ComVisible(true)]
    [ProgId("TestExcel2007Addin.SimpleControl")]
    [Guid("172E76F2-BC2E-4629-9980-DB2D8724EA73")]
    public partial class SimpleControl : UserControl
    {
        public SimpleControl()
        {
            InitializeComponent();
            label.Text += AppDomain.CurrentDomain.FriendlyName;
        }
    }
}

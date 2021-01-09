using System;
using System.Windows.Forms;
using System.Drawing;

namespace TestOutlook2007Addin
{
    internal class PictureConverter : AxHost
    {
        private PictureConverter() : base("") { }

        static public stdole.IPictureDisp IconToPictureDisp(Icon icon)
        {
            return (stdole.IPictureDisp)
                GetIPictureDispFromPicture(icon.ToBitmap());
        }
    }
}

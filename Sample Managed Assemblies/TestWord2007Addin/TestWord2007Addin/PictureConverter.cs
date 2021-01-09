using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Drawing;

namespace TestWord2007Addin
{
    internal class PictureConverter : AxHost
    {
        private PictureConverter() : base("") { }

        static public stdole.IPictureDisp ImageToPictureDisp(Image image)
        {
            return (stdole.IPictureDisp)GetIPictureDispFromPicture(image);
        }

        static public stdole.IPictureDisp IconToPictureDisp(Icon icon)
        {
            return ImageToPictureDisp(icon.ToBitmap());
        }
    }
}

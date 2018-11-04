using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using ApplicationClass = Microsoft.Office.Interop.Word.Application;
using Microsoft.Office.Tools.Ribbon;
using System.Drawing.Imaging;

namespace WordAddInResizeImage
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void Test_Click(object sender, RibbonControlEventArgs e)
        {
            foreach (InlineShape shape in Globals.ThisAddIn.Application.ActiveDocument.InlineShapes)
            {
                if (shape.Type == WdInlineShapeType.wdInlineShapePicture)
                {
                    shape.Select();
                    Selection selection = Globals.ThisAddIn.Application.Selection;
                    if (selection.Type == WdSelectionType.wdSelectionInlineShape)
                    {
                        //selection.CopyAsPicture();
                        //if (Clipboard.ContainsImage())
                        //{
                        //    System.Drawing.Image img = Clipboard.GetImage();
                        //    img.Save("d:\\temp\\temp" + count.ToString() + ".jpg", ImageFormat.Jpeg);
                        //}
                        string result = string.Format("width = {0}, height = {1} ", shape.Width, shape.Height);
                        selection.InsertAfter(result);
                    }

                    System.Threading.Thread.Sleep(2000);
                }
            }
        }
    }
}

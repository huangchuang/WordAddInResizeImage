using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddInResizeImage
{
    public class MyBookMark
    {
        public Word.Range BookMarkRange { get; set; }
        public string ToolTip { get; set; }
        public bool IsHighLighted { get; set; }
        public Word.WdColorIndex OrignalColor { get; set; }
        public Word.Bookmark BookMark { get; set; }
    }
}

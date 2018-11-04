using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Drawing;
using Microsoft.Office.Core;
using System.Windows.Forms;

namespace WordAddInResizeImage
{
    public partial class ThisAddIn
    {
        public Microsoft.Office.Tools.CustomTaskPane _MyCustomTaskPane = null;
        public List<MyBookMark> _BookMarkList = new List<MyBookMark>();
        public FloatingPanel _FloatingPanel = null;
        private CommandBar textCommandBar = null;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            textCommandBar = Application.CommandBars["Text"];

            AddCustomPane();

            RemoveRightClickMenu();
            AddRightClickMenu();

            RegisterApplicationEvents();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            RemoveRightClickMenu();
            UnregisterApplicationEvents();
        }

        private void RegisterApplicationEvents()
        {
            this.Application.WindowSelectionChange += new Word.ApplicationEvents4_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);
            this.Application.WindowBeforeRightClick += new Word.ApplicationEvents4_WindowBeforeRightClickEventHandler(Application_WindowBeforeRightClick);
        }

        private void UnregisterApplicationEvents()
        {
            this.Application.WindowSelectionChange -= new Word.ApplicationEvents4_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);
            this.Application.WindowBeforeRightClick -= new Word.ApplicationEvents4_WindowBeforeRightClickEventHandler(Application_WindowBeforeRightClick);
        }

        private void AddRightClickMenu()
        {
            // 添加右键按钮
            CommandBarControls controls = textCommandBar.Controls;
            CommandBarButton addBtn = (CommandBarButton)controls.Add(Office.MsoControlType.msoControlButton, missing, missing, missing, true);
            addBtn.BeginGroup = true;
            addBtn.Tag = "BookMarkAddin";
            addBtn.Caption = "Add Bookmark";
            addBtn.Enabled = true;
            addBtn.Click += new Office._CommandBarButtonEvents_ClickEventHandler(RightClickMenuHandler);
        }

        private void RemoveRightClickMenu()
        {
            // 这里我写了一个循环，目标是清理所有由我创建的右键按钮，尤其是由于Addin Crash时所遗留的按钮
            CommandBarControl control = textCommandBar.FindControl(MsoControlType.msoControlButton, missing, "BookMarkAddin", true, true);
            while (control != null)
            {
                control.Delete(true);
                control = textCommandBar.FindControl(MsoControlType.msoControlButton, missing, "BookMarkAddin", true, true);
            }
        }

        private void AddCustomPane()
        {
            TaskPane1 taskPane = new TaskPane1();
            _MyCustomTaskPane = this.CustomTaskPanes.Add(taskPane, "My Task Pane");
            _MyCustomTaskPane.Width = 200;
            _MyCustomTaskPane.Visible = true;
        }

        private void Application_WindowBeforeRightClick(Word.Selection Sel, ref bool Cancel)
        {
            if (!string.IsNullOrWhiteSpace(Sel.Range.Text) && Sel.Range.Text.Length > 2)
            {
                
            }
        }
        
        private void Application_WindowSelectionChange(Word.Selection Sel)
        {
            MyBookMark bookmark = _BookMarkList.FirstOrDefault(mark => Sel.Range.Start >= mark.BookMarkRange.Start && Sel.Range.Start <= mark.BookMarkRange.End);
            //if (bookmark != null)
            //{
            //    Globals.ThisAddIn._FloatingPanel = new FloatingPanel(bookmark);

            //    Point currentPos = GetPositionForShowing(Sel);

            //    Globals.ThisAddIn._FloatingPanel.Location = currentPos;
            //    Globals.ThisAddIn._FloatingPanel.Show();
            //}
        }

        private void RightClickMenuHandler(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            //Point currentPos = GetPositionForShowing(this.Application.Selection);
            System.Windows.Forms.Form form = new System.Windows.Forms.Form();
            //form.Location = currentPos;

            form.ShowDialog();
        }

        private static Point GetPositionForShowing(Word.Selection Sel)
        {
            // get range postion
            int left = 0;
            int top = 0;
            int width = 0;
            int height = 0;
            Globals.ThisAddIn.Application.ActiveDocument.ActiveWindow.GetPoint(out left, out top, out width, out height, Sel.Range);

            Point currentPos = new Point(left, top);
            if (Screen.PrimaryScreen.Bounds.Height - top > 340)
            {
                currentPos.Y += 20;
            }
            else
            {
                currentPos.Y -= 320;
            }
            return currentPos;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}

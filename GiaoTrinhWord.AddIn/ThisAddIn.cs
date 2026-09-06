using System;
using Microsoft.Office.Tools;

namespace GiaoTrinhWord.AddIn
{
    public partial class ThisAddIn
    {
        private TaskPaneControl taskPaneControl;
        private CustomTaskPane taskPane;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            taskPaneControl = new TaskPaneControl();

            taskPane = this.CustomTaskPanes.Add(
                taskPaneControl,
                "Giáo Trình Word"
            );

            taskPane.Width = 420;
            taskPane.Visible = false;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
        }

        public void ShowTaskPane()
        {
            if (taskPane != null)
            {
                taskPane.Visible = true;
            }
        }

        public void HideTaskPane()
        {
            if (taskPane != null)
            {
                taskPane.Visible = false;
            }
        }

        public void ToggleTaskPane()
        {
            if (taskPane != null)
            {
                taskPane.Visible = !taskPane.Visible;
            }
        }

        #region VSTO generated code

        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
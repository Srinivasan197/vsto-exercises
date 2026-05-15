using System;
using Microsoft.Office.Tools;


namespace Exercise_2_Add_a_Rectangle_to_Current_Slide
{
    public partial class ThisAddIn
    {
        // Global Task Pane reference
        public CustomTaskPane CustomTaskPane;
        private void ThisAddIn_Startup(object sender,EventArgs e)
        {
            // Create Task Pane UserControl
            TaskPaneControl taskPaneControl = new TaskPaneControl();

            // Add Task Pane
            CustomTaskPane = this.CustomTaskPanes.Add(taskPaneControl,"Slide Information");

            // Initially hidden
            CustomTaskPane.Visible = false;
        }

        private void ThisAddIn_Shutdown(object sender,EventArgs e)
        {

        }

        #region VSTO generated code

        private void InternalStartup()
        {
            this.Startup +=
                new EventHandler(ThisAddIn_Startup);

            this.Shutdown +=
                new EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
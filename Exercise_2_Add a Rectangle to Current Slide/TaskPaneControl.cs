using Exercise_5_Basic_Task_Pane.Service;
using System;
using System.Data.SqlClient;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Exercise_2_Add_a_Rectangle_to_Current_Slide
{
    public partial class TaskPaneControl : UserControl
    {
        public TaskPaneControl()
        {
            InitializeComponent();

            LoadSlideInformation();
        }

        private void LoadSlideInformation()
        {
            try
            {
                PowerPoint.Application application =
                    Globals.ThisAddIn.Application;

                PowerPoint.Slide currentSlide =
                    application.ActiveWindow.View.Slide;

                SlideInfoService slideInfoService =
                    new SlideInfoService();

                var result =
                    slideInfoService.GetCurrentSlideInformation(
                        currentSlide);

                if (!result.Success)
                {
                    lblError.Text = result.Message;

                    txtSlideInfo.Text = string.Empty;

                    return;
                }

                lblError.Text = string.Empty;

                txtSlideInfo.Text = result.Message;
            }
            catch (Exception ex)
            {
                lblError.Text = ex.Message;
            }
        }

        private void btnRefresh_Click(
            object sender,
            EventArgs e)
        {
            LoadSlideInformation();
        }

        private void btnClose_Click(
            object sender,
            EventArgs e)
        {
            Globals.ThisAddIn.CustomTaskPane.Visible =
                false;
        }
    }
}
using ExampleVSTO.CommonUtilities.Utility;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Windows.Forms;
using Exercise_2_Add_a_Rectangle_to_Current_Slide.Service;

namespace Exercise_2_Add_a_Rectangle_to_Current_Slide
{
    public partial class MyRibbon
    {
        private readonly Logger _logger = new Logger(typeof(MyRibbon));
        private void MyRibbon_Load(object sender,RibbonUIEventArgs e)
        {
            _logger.Info("MyRibbon loaded successfully.");
        }
        private void AddRectangular_Click(object sender , RibbonControlEventArgs e)
        {
            _logger.Info("Add Rectangle button clicked.");

            try
            {
                RectangleService rectangleService = new RectangleService();

                BooleanResult<string> result = rectangleService.AddRectangleToCurrentSlide();

                if (!result.Success)
                {
                    _logger.Warn(
                        result.Message);

                    MessageBox.Show(
                        result.Message,
                        "Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);

                    return;
                }
                MessageBox.Show(
                    result.Message,
                    "Success",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                _logger.Info("Rectangle creation completed.");
            }
            catch (Exception ex)
            {
                _logger.Error(
                    "Unexpected error occurred in ribbon.",
                    ex);

                MessageBox.Show(
                    "Unexpected error occurred.",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
    }
}
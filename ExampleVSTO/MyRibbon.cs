using ExampleVSTO.CommonUtilities.Utility;
using ExampleVSTO.Service;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;

namespace ExampleVSTO
{
    public partial class MyRibbon
    {
        private readonly Logger _logger = new Logger(typeof(MyRibbon));
        /// Ribbon Load Event
        private void MyRibbon_Load(object sender,RibbonUIEventArgs e)
        {
            _logger.Info("MyRibbon loaded successfully.");
        }
        /// Button Click Event
        private void button1_Click(object sender,RibbonControlEventArgs e)
        {
            _logger.Info("Presentation Info button clicked.");

            PresentationReader reader =new PresentationReader();

            BooleanResult<string> result =reader.GetPresentationInformation();

            // Failure handling
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

            // Success
            MessageBox.Show(
                result.Result,
                "Presentation Info",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);

            _logger.Info("Presentation information displayed.");
        }
    }
}
using ExampleVSTO.CommonUtilities.Utility;
using Exercise_2_Add_a_Rectangle_to_Current_Slide.Service;
using Exercise_4_Group_Shape_Traversal.Service;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

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
                    "Rectangle added successfully.",
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
        private void btnEnumerateShapes_Click(object sender,RibbonControlEventArgs e)
        {
            _logger.Info("Enumerate Shapes button clicked.");

            try
            {             
                ShapeDiagnosticService shapeDiagnosticService = new ShapeDiagnosticService();

                BooleanResult<string> result = shapeDiagnosticService.ExportShapeDiagnostics();

                if (!result.Success)
                {
                    _logger.Warn(result.Message);

                    MessageBox.Show(
                    "Enumerate Shapes details done successfully.",
                    "Success",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                    return;
                }
                MessageBox.Show(
                    result.Message,
                    "Success",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                _logger.Info("Shape enumeration completed.");
            }
            catch (Exception ex)
            {
                _logger.Error(
                    "Unexpected error occurred while enumerating shapes.",
                    ex);

                MessageBox.Show(
                    "Unexpected error occurred.",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
        private void ShapeTraversal_Click(object sender,RibbonControlEventArgs e)
        {
            _logger.Info("Shape Traversal button clicked.");

            try
            {
                PowerPoint.Application application = Globals.ThisAddIn.Application;

                PowerPoint.Slide currentSlide = application.ActiveWindow.View.Slide;

                ShapeTraversalService shapeTraversalService = new ShapeTraversalService();

                BooleanResult<string> result = shapeTraversalService.GenerateShapeReport(currentSlide);

                if (!result.Success)
                {
                    _logger.Warn(result.Message);

                    MessageBox.Show(
                        result.Message,
                        "Warning",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);

                    return;
                }

                MessageBox.Show(
                    result.Message,
                    "Success",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                _logger.Info("Shape traversal completed successfully.");
            }
            catch (Exception ex)
            {
                _logger.Error(
                    "Unexpected error occurred while traversing shapes.",
                    ex);

                MessageBox.Show(
                    "Unexpected error occurred.",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void btnShowTaskPane_Click(object sender, RibbonControlEventArgs e)
        {
            if (Globals.ThisAddIn.CustomTaskPane != null)
            {
                Globals.ThisAddIn.CustomTaskPane.Visible = !Globals.ThisAddIn.CustomTaskPane.Visible;
            }
        }
    }
}
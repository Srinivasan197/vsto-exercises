using ExampleVSTO.CommonUtilities.Utility;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Text;
using System.Windows.Forms;

namespace Exercise_2_Add_a_Rectangle_to_Current_Slide.Service
{
    public class ShapeDiagnosticService
    {
        private readonly Logger _logger = new Logger(typeof(ShapeDiagnosticService));
        public BooleanResult<string> ExportShapeDiagnostics()
        {
            try
            {
                _logger.Info("Starting shape diagnostics export.");
                Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
                if (presentation == null)
                {
                    _logger.Warn("No active presentation found.");
                    return BooleanResult<string>.FailResult("No active presentation found.");
                }
                StringBuilder report = new StringBuilder();
                foreach (Slide slide in presentation.Slides)
                {
                    _logger.Info($"Processing Slide {slide.SlideIndex}");
                    report.AppendLine($"Slide {slide.SlideIndex}");
                    foreach (Shape shape in slide.Shapes)
                    {
                        _logger.Info($"Reading Shape : {shape.Name}");
                        report.AppendLine($"Shape Name : {shape.Name}");
                        report.AppendLine($"Shape Type : {shape.Type}");
                        report.AppendLine($"Left       : {shape.Left}");
                        report.AppendLine($"Top        : {shape.Top}");
                        report.AppendLine($"Width      : {shape.Width}");
                        report.AppendLine($"Height     : {shape.Height}");
                        // Has Text Frame
                        bool hasTextFrame = shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue;
                        report.AppendLine($"Has Text Frame : {hasTextFrame}");
                        // Is Group
                        bool isGroup = shape.Type == Microsoft.Office.Core.MsoShapeType.msoGroup;
                        report.AppendLine($"Is Group : {isGroup}");
                        // Placeholder Type
                        if (shape.PlaceholderFormat != null)
                        {
                            try
                            {
                                report.AppendLine(
                                    $"Placeholder Type : {shape.PlaceholderFormat.Type}");
                            }
                            catch
                            {
                                report.AppendLine("Placeholder Type : Not Available");
                            }
                        }
                        report.AppendLine("--------");
                    }
                }
                // Show report
                MessageBox.Show(report.ToString());
                _logger.Info("Shape diagnostics exported successfully.");
                return BooleanResult<string>.SuccessResult("Shape diagnostics exported successfully.");
            }
            catch (Exception ex)
            {
                _logger.Error("Error while exporting shape diagnostics.", ex);
                MessageBox.Show("Failed to export shape diagnostics.\n" + ex.Message);
                return BooleanResult<string>.FailResult($"Failed to export shape diagnostics: {ex.Message}");
            }
        }
    }
}
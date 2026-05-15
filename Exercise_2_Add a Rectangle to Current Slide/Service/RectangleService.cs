using ExampleVSTO.CommonUtilities.Utility;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System;

namespace Exercise_2_Add_a_Rectangle_to_Current_Slide.Service
{
    public class RectangleService
    {
        private readonly Logger _logger = new Logger(typeof(RectangleService));
        public BooleanResult<string> AddRectangleToCurrentSlide()
        {
            _logger.Info("AddRectangleToCurrentSlide started.");
            try
            {
                Application app = Globals.ThisAddIn.Application;

                if (app == null)
                {
                    return BooleanResult<string>.FailResult("PowerPoint application not available.");
                }
                _logger.Info("PowerPoint application validated.");

                if (app.Presentations.Count == 0)
                {
                    return BooleanResult<string>.FailResult("No active presentation found.");
                }
                Presentation presentation = app.ActivePresentation;
                if (presentation == null)
                {
                    return BooleanResult<string>.FailResult("Active presentation is null.");
                }
                _logger.Info($"Presentation validated: {presentation.Name}");
                DocumentWindow window = app.ActiveWindow;
                if (window == null)
                {
                    return BooleanResult<string>.FailResult("No active PowerPoint window found.");
                }
                _logger.Info("Active window validated.");
                Slide slide = null;
                try
                {
                    slide = window.View?.Slide;
                }
                catch (Exception ex)
                {
                    _logger.Error("Failed to access active slide.",ex);

                    return BooleanResult<string>.FailResult("Unable to access active slide.");
                }
                if (slide == null)
                {
                    return BooleanResult<string>.FailResult("No active slide found.");
                }
                _logger.Info($"Active slide validated. Slide Index: {slide.SlideIndex}");
                Shape rectangle = null;
                try
                {
                    _logger.Info("Adding rectangle shape.");

                    rectangle = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle,
                        100,
                        100,
                        300,
                        100);
                }
                catch (Exception ex)
                {
                    _logger.Error("Failed to add rectangle shape.",ex);

                    return BooleanResult<string>.FailResult("Failed to add rectangle.");
                }
                if (rectangle == null)
                {
                    return BooleanResult<string>.FailResult("Rectangle creation returned null.");
                }
                _logger.Info($"Rectangle created successfully: {rectangle.Name}");
                ApplyFillColor(rectangle);
                ApplyBorder(rectangle);
                ApplyText(rectangle);
                _logger.Info("Rectangle creation completed successfully.");
                return BooleanResult<string>.SuccessResult("Rectangle added successfully.");
            }
            catch (Exception ex)
            {
                _logger.Error("Unexpected error occurred in RectangleService.",ex);
                return BooleanResult<string>.FailResult($"Unexpected error: {ex.Message}");
            }
        }
        private void ApplyFillColor(Shape rectangle)
        {
            try
            {
                _logger.Info("Applying fill color.");
                rectangle.Fill.ForeColor.RGB = System.Drawing.Color.LightBlue.ToArgb();
            }
            catch (Exception ex)
            {
                _logger.Error("Failed to apply fill color.",ex);
                throw;
            }
        }
        private void ApplyBorder(Shape rectangle)
        {
            try
            {
                _logger.Info("Applying border.");
                rectangle.Line.ForeColor.RGB = System.Drawing.Color.DarkBlue.ToArgb();
                rectangle.Line.Weight = 2;
            }
            catch (Exception ex)
            {
                _logger.Error("Failed to apply border.",ex);
                throw;
            }
        }
        private void ApplyText(Shape rectangle)
        {
            try
            {
                _logger.Info("Applying text.");

                if (rectangle.HasTextFrame ==
                    Office.MsoTriState.msoTrue)
                {
                    rectangle.TextFrame.TextRange.Text ="Hello from VSTO Add-in";
                    rectangle.TextFrame.TextRange.Font.Size = 20;
                    rectangle.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                }
                else
                {
                    _logger.Warn("Rectangle does not support text frame.");
                }
            }
            catch (Exception ex)
            {
                _logger.Error("Failed to apply text.",ex);
                throw;
            }
        }
    }
}
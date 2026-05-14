using ExampleVSTO.CommonUtilities.Utility;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Text;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace ExampleVSTO
{
    public partial class MyRibbon
    {
        // Logger instance
        private readonly Logger _logger = new Logger(typeof(MyRibbon));

        /// <summary>
        /// Ribbon Load Event
        /// </summary>
        private void MyRibbon_Load(object sender,
            RibbonUIEventArgs e)
        {
            _logger.Info("MyRibbon loaded successfully.");
        }

        /// <summary>
        /// Button Click Event
        /// </summary>
        private void button1_Click(object sender,
            RibbonControlEventArgs e)
        {
            _logger.Info("Presentation Info button clicked.");

            BooleanResult<string> result =  GetPresentationInformation();

            // Failure handling
            if (!result.Success)
            {
                _logger.Warn(result.Message);

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

        /// <summary>
        /// Reads PowerPoint presentation information safely
        /// </summary>
        private BooleanResult<string> GetPresentationInformation()
        {
            _logger.Info("GetPresentationInformation started.");

            try
            {
                // Get PowerPoint Application
                PowerPoint.Application app =
                    Globals.ThisAddIn.Application;

                // Validate application
                if (app == null)
                {
                    return BooleanResult<string>
                        .FailResult(
                            "PowerPoint application is not available.");
                }

                _logger.Info("PowerPoint application validated.");

                // Get active presentation
                PowerPoint.Presentation pres =
                    app.ActivePresentation;

                // Validate presentation
                if (pres == null)
                {
                    return BooleanResult<string>
                        .FailResult(
                            "No active presentation found.");
                }

                _logger.Info(
                    $"Active presentation found: {pres.Name}");

                // Get active window
                PowerPoint.DocumentWindow window =
                    app.ActiveWindow;

                // Validate window
                if (window == null)
                {
                    return BooleanResult<string>
                        .FailResult(
                            "No active PowerPoint window found.");
                }

                _logger.Info("Active window validated.");

                // Get active slide safely
                PowerPoint.Slide slide = null;

                try
                {
                    slide = window.View?.Slide;
                }
                catch (Exception ex)
                {
                    _logger.Error(
                        "Failed to access active slide.",
                        ex);

                    return BooleanResult<string>
                        .FailResult(
                            "Unable to access active slide.");
                }

                // Validate slide
                if (slide == null)
                {
                    return BooleanResult<string>
                        .FailResult(
                            "No active slide found.");
                }
                _logger.Info(
                    $"Active slide index: {slide.SlideIndex}");

                // Get selection safely
                PowerPoint.Selection selection = null;

                try
                {
                    selection = window.Selection;
                }
                catch (Exception ex)
                {
                    _logger.Error(
                        "Failed to access selection.",
                        ex);

                    return BooleanResult<string>
                        .FailResult(
                            "Unable to access selection.");
                }

                // Build output safely
                StringBuilder info =
                    new StringBuilder();

                info.AppendLine(
                    $"Presentation Name: {pres.Name}");

                info.AppendLine(
                    $"Total Slides: {pres.Slides.Count}");

                info.AppendLine(
                    $"Current Slide Number: {slide.SlideIndex}");

                // Validate selection
                if (selection == null)
                {
                    info.AppendLine(
                        "No active selection found.");

                    _logger.Warn(
                        "Selection object is null.");
                }
                else if (selection.Type ==
                    PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    _logger.Info(
                        "Shape selection detected.");

                    int shapeCount = 0;

                    try
                    {
                        shapeCount =
                            selection.ShapeRange.Count;
                    }
                    catch (Exception ex)
                    {
                        _logger.Error(
                            "Failed to read ShapeRange count.",
                            ex);

                        return BooleanResult<string> .FailResult("Unable to read selected shapes.");
                    }

                    info.AppendLine(
                        $"Selected Shapes Count: {shapeCount}");

                    // Traverse shapes safely
                    for (int i = 1; i <= shapeCount; i++)
                    {
                        try
                        {
                            PowerPoint.Shape shape =
                                selection.ShapeRange[i];

                            if (shape == null)
                            {
                                info.AppendLine(
                                    $"Shape {i}: NULL");

                                continue;
                            }

                            info.AppendLine(
                                $"Shape {i}: {shape.Name} ({shape.Type})");

                            _logger.Info(
                                $"Shape {i} read successfully.");

                            // Detect group shapes
                            if (shape.Type ==
                                Office.MsoShapeType.msoGroup)
                            {
                                info.AppendLine(
                                    $"   -> Group Shape Detected");

                                try
                                {
                                    PowerPoint.GroupShapes groupShapes =
                                        shape.GroupItems;

                                    info.AppendLine(
                                        $"   -> Child Shape Count: {groupShapes.Count}");

                                    for (int j = 1;
                                         j <= groupShapes.Count;
                                         j++)
                                    {
                                        PowerPoint.Shape childShape =
                                            groupShapes[j];

                                        info.AppendLine(
                                            $"      Child {j}: {childShape.Name}");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    _logger.Error(
                                        "Failed to read group shapes.",
                                        ex);

                                    info.AppendLine(
                                        "   -> Failed to read child group shapes.");
                                }
                            }
                            // Validate text frame safely
                            if (shape.HasTextFrame ==
                                Office.MsoTriState.msoTrue)
                            {
                                if (shape.TextFrame.HasText ==
                                    Office.MsoTriState.msoTrue)
                                {
                                    string text =
                                        shape.TextFrame.TextRange.Text;

                                    info.AppendLine(
                                        $"   -> Text: {text}");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger.Error(
                                $"Failed to process shape index {i}.",
                                ex);

                            info.AppendLine(
                                $"Shape {i}: Failed to process.");
                        }
                    }
                }
                else
                {
                    info.AppendLine(
                        "No shapes selected.");

                    _logger.Info(
                        $"Selection type is: {selection.Type}");
                }

                _logger.Info(
                    "GetPresentationInformation completed successfully.");

                return BooleanResult<string>
                    .SuccessResult(info.ToString());
            }
            catch (Exception ex)
            {
                _logger.Error(
                    "Unexpected error in GetPresentationInformation.",
                    ex);

                return BooleanResult<string>
                    .FailResult(
                        "Unexpected error occurred: " + ex.Message);
            }
        }
    }
}
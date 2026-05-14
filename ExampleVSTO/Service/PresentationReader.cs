using ExampleVSTO.CommonUtilities.Utility;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System;
using System.Text;

namespace ExampleVSTO.Service
{
    public class PresentationReader
    {
        private readonly Logger _logger = new Logger(typeof(PresentationReader));
        public BooleanResult<string> GetPresentationInformation()
        {
            _logger.Info("GetPresentationInformation started.");
            try
            {
                // Get PowerPoint Application
                Application app = Globals.ThisAddIn.Application;

                // Validate application
                if (app == null)
                {
                    return BooleanResult<string>.FailResult("PowerPoint application is not available.");
                }
                _logger.Info("PowerPoint application validated.");

                Presentation pres = app.ActivePresentation;

                if (pres == null)
                {
                    return BooleanResult<string>.FailResult("No active presentation found.");
                }

                _logger.Info($"Active presentation found: {pres.Name}");

                // Get active window
                DocumentWindow window = app.ActiveWindow;

                // Validate window
                if (window == null)
                {
                    return BooleanResult<string>.FailResult("No active PowerPoint window found.");
                }
                _logger.Info("Active window validated.");

                // Get active slide safely
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
                // Validate slide
                if (slide == null)
                {
                    return BooleanResult<string>.FailResult("No active slide found.");
                }
                _logger.Info($"Active slide index: {slide.SlideIndex}");
                // Get selection safely
                Selection selection = null;
                try
                {
                    selection = window.Selection;
                }
                catch (Exception ex)
                {
                    _logger.Error("Failed to access selection.",ex);
                    return BooleanResult<string>.FailResult("Unable to access selection.");
                }

                // Build output
                StringBuilder info =new StringBuilder();

                info.AppendLine($"Presentation Name: {pres.Name}");

                info.AppendLine($"Total Slides: {pres.Slides.Count}");

                info.AppendLine($"Current Slide Number: {slide.SlideIndex}");

                // Read selection information
                ReadSelectionInformation(selection,info);

                _logger.Info("GetPresentationInformation completed successfully.");

                return BooleanResult<string>.SuccessResult(info.ToString());
            }
            catch (Exception ex)
            {
                _logger.Error("Unexpected error occurred.",ex);

                return BooleanResult<string>.FailResult("Unexpected error occurred: "+ ex.Message);
            }
        }
        /// Reads selection information safely.
        private void ReadSelectionInformation(Selection selection,StringBuilder info)
        {
            if (selection == null)
            {
                info.AppendLine(
                    "No active selection found.");

                _logger.Warn(
                    "Selection object is null.");

                return;
            }
            // Shape selection
            if (selection.Type == PpSelectionType.ppSelectionShapes)
            {
                _logger.Info("Shape selection detected.");

                ShapeRange shapeRange = selection.ShapeRange;

                int shapeCount = shapeRange.Count;

                info.AppendLine($"Selected Shapes Count: {shapeCount}");

                for (int i = 1; i <= shapeCount; i++)
                {
                    ReadShapeInformation(shapeRange[i],info,i);
                }
            }
            else
            {
                info.AppendLine("No shapes selected.");

                _logger.Info($"Selection type: {selection.Type}");
            }
        }
        /// Reads shape information safely.
        private void ReadShapeInformation(Shape shape,StringBuilder info,int index)
        {
            try
            {
                if (shape == null)
                {
                    info.AppendLine($"Shape {index}: NULL");
                    return;
                }
                info.AppendLine($"Shape {index}:{shape.Name}({shape.Type})");
                _logger.Info($"Shape {index} processed.");
                // Group shape handling
                if (shape.Type ==
                    Office.MsoShapeType.msoGroup)
                {
                    ReadGroupShapes(shape,info);
                }
                // Read text safely
                ReadShapeText(shape,info);
            }
            catch (Exception ex)
            {
                _logger.Error($"Failed processing shape {index}.",ex);

                info.AppendLine($"Shape {index}: Failed to process.");
            }
        }
        /// Reads group shapes safely.
        private void ReadGroupShapes(Shape shape,StringBuilder info)
        {
            try
            {
                GroupShapes groupShapes = shape.GroupItems;

                info.AppendLine($"Child Shape Count: {groupShapes.Count}");

                for (int i = 1;i <= groupShapes.Count;i++)
                {
                    Shape childShape = groupShapes[i];

                    info.AppendLine($"Child {i}: {childShape.Name}");
                }
            }
            catch (Exception ex)
            {
                _logger.Error("Failed reading group shapes.",ex);
                info.AppendLine("Failed reading child shapes.");
            }
        }
        /// Reads shape text safely.
        private void ReadShapeText(Shape shape,StringBuilder info)
        {
            try
            {
                if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                {
                    if (shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                    {
                        string text = shape.TextFrame.TextRange.Text;

                        info.AppendLine($"Text: {text}");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error("Failed reading shape text.",ex);
            }
        }
    }
}
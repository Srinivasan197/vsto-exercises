using ExampleVSTO.CommonUtilities.Utility;
using Microsoft.Office.Interop.PowerPoint;
using System;

namespace Exercise_5_Basic_Task_Pane.Service
{
    public class SlideInfoService
    {
        private readonly Logger _logger = new Logger(typeof(SlideInfoService));
        public BooleanResult<string> GetCurrentSlideInformation(Slide currentSlide)
        {
            try
            {
                _logger.Info("Loading current slide information.");
                if (currentSlide == null)
                {
                    return new BooleanResult<string>
                    {
                        Success = false,
                        Message = "No active slide found."
                    };
                }
                string slideInformation =
                    $"Slide Index : {currentSlide.SlideIndex}" +
                    Environment.NewLine +
                    $"Slide Name : {currentSlide.Name}" +
                    Environment.NewLine +
                    $"Total Shapes : {currentSlide.Shapes.Count}";
                _logger.Info("Slide information loaded successfully.");
                return new BooleanResult<string>
                {
                    Success = true,
                    Message = slideInformation
                };
            }
            catch (Exception ex)
            {
                _logger.Error("Error while loading slide information.",ex);
                return new BooleanResult<string>
                {
                    Success = false,
                    Message = ex.Message
                };
            }
        }
    }
}
using ExampleVSTO.CommonUtilities.Utility;
using Exercise_2_Add_a_Rectangle_to_Current_Slide.Service;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Text;
using Office = Microsoft.Office.Core;

namespace Exercise_4_Group_Shape_Traversal.Service
{
    public class ShapeTraversalService
    {
        private readonly Logger _logger = new Logger(typeof(ShapeDiagnosticService));
        private readonly StringBuilder _report = new StringBuilder();
        public BooleanResult<string> GenerateShapeReport(Slide slide)
        {
            try
            {
                _logger.Info("Starting shape traversal.");
                _report.Clear();
                TraverseShapes(slide.Shapes, 0);
                _logger.Info("Shape traversal completed successfully.");
                return new BooleanResult<string>
                {
                    Success = true,
                    Message = _report.ToString()
                };
            }
            catch (Exception ex)
            {
                _logger.Error("Error occurred during shape traversal.",ex);
                return new BooleanResult<string>
                {
                    Success = false,
                    Message = ex.Message
                };
            }
        }
        private void TraverseShapes(dynamic shapes , int level)
        {
            foreach (Shape shape in shapes)
            {
                _logger.Info($"Traversing Shape : {shape.Name}");
                string indent = new string(' ', level * 4);
                _report.AppendLine($"{indent}Shape Name : {shape.Name}");
                _report.AppendLine($"{indent}Shape Type : {shape.Type}");
                _report.AppendLine();

                if (shape.Type == Office.MsoShapeType.msoGroup)
                {
                    _report.AppendLine($"{indent}Entering Group : {shape.Name}");
                    _report.AppendLine();
                    TraverseShapes(shape.GroupItems,level + 1);
                }
            }
        }
    }
}
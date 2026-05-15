using ExampleVSTO.CommonUtilities.Utility;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System;
using System.Text;

namespace Exercise_4_Group_Shape_Traversal.Service
{
    public class ShapeTraversalService
    {
        private readonly StringBuilder _report = new StringBuilder();
        public BooleanResult<string> GenerateShapeReport(Slide slide)
        {
            try
            {
                _report.Clear();
                TraverseShapes(slide.Shapes, 0);
                return new BooleanResult<string>
                {
                    Success = true,
                    Message = _report.ToString()
                };
            }
            catch (Exception ex)
            {
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
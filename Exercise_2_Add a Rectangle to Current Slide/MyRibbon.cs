using Microsoft.Office.Tools.Ribbon;
using System;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace Exercise_2_Add_a_Rectangle_to_Current_Slide
{
    public partial class MyRibbon
    {
        private void MyRibbon_Load(object sender,
    RibbonUIEventArgs e)
        {
           // _logger.Info("MyRibbon loaded successfully.");
        }
        private void AddRectangular_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                MessageBox.Show("Starting rectangle creation...");

                PowerPoint.Application app = Globals.ThisAddIn.Application;

                if (app.Presentations.Count == 0)
                {
                    MessageBox.Show("No active presentation found.");
                    return;
                }
                MessageBox.Show("Presentation validated.");

                PowerPoint.DocumentWindow window = app.ActiveWindow;

                if (window == null || window.View.Slide == null)
                {
                    MessageBox.Show("No active slide found.");
                    return;
                }
                PowerPoint.Slide slide = window.View.Slide;
                MessageBox.Show("Active slide validated.");

                PowerPoint.Shape rect = slide.Shapes.AddShape(
                        Office.MsoAutoShapeType.msoShapeRectangle,
                        100,   
                        100,   
                        300,   
                        100    
                    );
                MessageBox.Show("Rectangle added.");
 
                rect.Fill.ForeColor.RGB = System.Drawing.Color.LightBlue.ToArgb();
                MessageBox.Show("Fill color applied.");

                rect.Line.ForeColor.RGB = System.Drawing.Color.DarkBlue.ToArgb();
                rect.Line.Weight = 2;
                MessageBox.Show("Border applied.");
  
                rect.TextFrame.TextRange.Text = "Hello from VSTO Add-in";

                rect.TextFrame.TextRange.Font.Size = 20;
                rect.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                MessageBox.Show("Text added.");

                MessageBox.Show("Rectangle creation completed successfully.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Error: " + ex.Message
                );
            }
        }
    }
}

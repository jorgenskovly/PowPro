using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowPro
{
    public partial class PowProRibbon : Office.IRibbonExtensibility
    {
        public void swap_onAction(Office.IRibbonControl control)
        {
            PowerPoint.ShapeRange selectedShapes = this.selectedShapes();

            if (selectedShapes.Count != 2) return;

            PowerPoint.Shape firstShape = selectedShapes[1];
            PowerPoint.Shape secondShape = selectedShapes[2];
            Point firstShapeCenter = ShapeController.getShapeCenter(firstShape);
            Point secondShapeCenter = ShapeController.getShapeCenter(secondShape);
            int firstShapeZOrder = firstShape.ZOrderPosition;
            int secondShapeZOrder = secondShape.ZOrderPosition;
            ShapeController.moveShapeCenter(firstShape, secondShapeCenter);
            ShapeController.moveShapeCenter(secondShape, firstShapeCenter);
            ShapeController.setShapeZOrder(firstShape, secondShapeZOrder);
            ShapeController.setShapeZOrder(secondShape, firstShapeZOrder);
        }
    }
}

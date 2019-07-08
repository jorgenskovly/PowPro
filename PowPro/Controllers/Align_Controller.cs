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
        public void sameRotation_onAction(Office.IRibbonControl control)
        {
            PowerPoint.ShapeRange shapes = this.selectedShapes();
            float rotation = shapes[1].Rotation;
            foreach (PowerPoint.Shape shape in shapes)
                shape.Rotation = rotation;
        }
        public void centerHorizontal_onAction(Office.IRibbonControl control)
        {
            PowerPoint.ShapeRange shapes = this.selectedShapes();
            float horizontal = ShapeController.getShapeHorizontalCenter(shapes[1]);
            foreach (PowerPoint.Shape shape in shapes)
                ShapeController.moveShapeHorizontalCenter(shape: shape, newShapeHorizontalCenter: horizontal);
        }

        public void centerVertical_onAction(Office.IRibbonControl control)
        {
            PowerPoint.ShapeRange shapes = this.selectedShapes();
            float verticalCenter = ShapeController.getShapeVerticalCenter(shapes[1]);
            foreach (PowerPoint.Shape shape in shapes)
                ShapeController.moveShapeVerticalCenter(shape: shape, newShapeVerticalCenter: verticalCenter);
        }
    }
}

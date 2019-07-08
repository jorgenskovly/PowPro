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
    public class ShapeController
    {
        public static Point getShapeCenter(PowerPoint.Shape myShape)
        {
            Point shapeCenter = new Point();
            shapeCenter.X = getShapeHorizontalCenter(myShape);
            shapeCenter.Y = getShapeVerticalCenter(myShape);
            return shapeCenter;
        }

        public static float getShapeHorizontalCenter(PowerPoint.Shape myShape)
        {
            return myShape.Left + (myShape.Width / 2);
        }
        public static float getShapeVerticalCenter(PowerPoint.Shape myShape)
        {
            return myShape.Top + (myShape.Height / 2);
        }

        public static void moveShapeCenter(PowerPoint.Shape shape, Point newShapeCenter)
        {
            moveShapeHorizontalCenter(shape, newShapeCenter.X);
            moveShapeVerticalCenter(shape, newShapeCenter.Y);
        }

        public static void moveShapeHorizontalCenter(PowerPoint.Shape shape, float newShapeHorizontalCenter)
        {
            shape.Left = newShapeHorizontalCenter - (shape.Width / 2);
        }

        public static void moveShapeVerticalCenter(PowerPoint.Shape shape, float newShapeVerticalCenter)
        {
            shape.Top = newShapeVerticalCenter - (shape.Height / 2);
        }

        public static void setShapeZOrder(PowerPoint.Shape shape, int newShapeZOrder)
        {            
            while (shape.ZOrderPosition < newShapeZOrder)
                shape.ZOrder(Office.MsoZOrderCmd.msoBringForward);
            while (shape.ZOrderPosition > newShapeZOrder)
                shape.ZOrder(Office.MsoZOrderCmd.msoSendBackward);
        }
    }
}

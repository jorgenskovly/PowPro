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
            //return myShape.ZOrderPosition;
            Point shapeCenter = new Point();
            shapeCenter.X = myShape.Left + (myShape.Width / 2);
            shapeCenter.Y = myShape.Top + (myShape.Height / 2);
            return shapeCenter;
        }

        internal static void moveShapeCenter(PowerPoint.Shape shape, Point newShapeCenter)
        {
            shape.Left = newShapeCenter.X - (shape.Width / 2);
            shape.Top = newShapeCenter.Y - (shape.Height / 2);
        }

        internal static void setShapeZOrder(PowerPoint.Shape shape, int newShapeZOrder)
        {
            //MessageBox.Show(shape.Name + ": move to ZOrder " + newShapeZOrder);
            
            while (shape.ZOrderPosition < newShapeZOrder)
                shape.ZOrder(Office.MsoZOrderCmd.msoBringForward);
            while (shape.ZOrderPosition > newShapeZOrder)
                shape.ZOrder(Office.MsoZOrderCmd.msoSendBackward);
        }
    }
}

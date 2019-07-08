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
        public void sameSize_onACtion(Office.IRibbonControl control)
        {
            sameSize(shapes: this.selectedShapes(), setHeight: true, setWidth: true);
        }
        public void sameHeight_onACtion(Office.IRibbonControl control)
        {
            sameSize(shapes: this.selectedShapes(), setHeight: true, setWidth: false);
        }
        public void sameWidth_onACtion(Office.IRibbonControl control)
        {
            sameSize(shapes: this.selectedShapes(), setHeight: false, setWidth: true);
        }

        public void sameSize(PowerPoint.ShapeRange shapes, bool setHeight, bool setWidth)
        {
            float height = shapes[1].Height;
            float width = shapes[1].Width;
            foreach (PowerPoint.Shape shape in shapes)
            {
                if (setHeight) shape.Height = height;
                if (setWidth) shape.Width = width;
            }
        }
    }
}

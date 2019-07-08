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
        public PowerPoint.ShapeRange selectedShapes()
        {
            PowerPoint.Application app = getCurrentApplication();
            if (app.ActiveWindow.Selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes) return null;
            else return app.ActiveWindow.Selection.HasChildShapeRange ? 
                    app.ActiveWindow.Selection.ChildShapeRange : app.ActiveWindow.Selection.ShapeRange;
        }

        public int selectedShapesCount()
        {
            PowerPoint.ShapeRange selectedShapes = this.selectedShapes();
            return selectedShapes == null ? 0 : selectedShapes.Count;
        }

        public bool exactlyTwoObjectSelected(Office.IRibbonControl control)
        {
            if (selectedShapesCount() == 2) return true;
            else return false;
        }

        public bool atLeastTwoObjectsSelected(Office.IRibbonControl control)
        {
            if (selectedShapesCount() >= 2) return true;
            else return false;
        }

        #region event handlers
        private void PowProRibbon_WindowSelectionChange(Microsoft.Office.Interop.PowerPoint.Selection Sel)
        {
            Console.WriteLine(DateTime.Now.ToString() +  " PowProRibbon_WindowSelectionChange()");
            this.ribbon.Invalidate();
        }

        private void PowProRibbon_SlideSelectionChanged(Microsoft.Office.Interop.PowerPoint.SlideRange SldRange)
        {
            return;
        }

        #endregion
    }
}

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
        public void onAction_callback(Office.IRibbonControl control)
        {
            return;
        }
    }
}

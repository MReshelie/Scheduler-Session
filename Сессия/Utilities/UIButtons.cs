using DevExpress.XtraBars.Docking2010;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Сессия
{
    public class UIButtons
    {
        static public bool UIButtonsEnabled(WindowsUIButtonPanel winUIBtPanel, bool flag)
        {
            winUIBtPanel.Buttons[1].Properties.Enabled = flag;
            winUIBtPanel.Buttons[3].Properties.Enabled = flag;
            return true;
        }
    }
}

using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace GiaoTrinhWord.AddIn
{
    public partial class RibbonGiaoTrinh
    {
        private void RibbonGiaoTrinh_Load(object sender, RibbonUIEventArgs e)
        {

        }

      

        private void btnMoGiaoTrinh_Click_1(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ShowTaskPane();
        }
    }
}

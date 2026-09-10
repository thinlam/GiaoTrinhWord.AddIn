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

       

        private void btnKiemTraChinhTa_Click_1(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;

                if (app.Documents.Count == 0)
                {
                    System.Windows.Forms.MessageBox.Show(
                        "Vui lòng mở hoặc tạo một tài liệu Word trước khi kiểm tra.",
                        "GIÁO TRÌNH WORD",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information
                    );

                    return;
                }

                app.ActiveDocument.CheckGrammar();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Không thể kiểm tra chính tả và ngữ pháp.\n\n" +
                    ex.Message,
                    "GIÁO TRÌNH WORD",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error
                );
            }
        }
    }
}

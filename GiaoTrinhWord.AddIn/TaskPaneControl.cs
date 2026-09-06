using System;
using System.IO;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace GiaoTrinhWord.AddIn
{
    public partial class TaskPaneControl : UserControl
    {
        private WebView2 webView;

        public TaskPaneControl()
        {
            InitializeComponent();

            InitializeWebViewControl();

            this.Load += TaskPaneControl_Load;
        }

        /// <summary>
        /// Khởi tạo WebView2 bằng code.
        /// Không tạo WebView2 trong Designer để tránh lỗi
        /// "already initialized with a different CoreWebView2Environment".
        /// </summary>
        private void InitializeWebViewControl()
        {
            webView = new WebView2
            {
                Dock = DockStyle.Fill,
                Name = "webViewGiaoTrinh"
            };

            Controls.Add(webView);
        }

        /// <summary>
        /// Khởi tạo môi trường WebView2 và mở website Giáo Trình Word.
        /// </summary>
        private async void TaskPaneControl_Load(object sender, EventArgs e)
        {
            try
            {
                // Lưu cache WebView2 ở thư mục người dùng.
                // Tránh lỗi E_ACCESSDENIED khi Word chạy từ Program Files.
                string userDataFolder = Path.Combine(
                    Environment.GetFolderPath(
                        Environment.SpecialFolder.LocalApplicationData
                    ),
                    "GiaoTrinhWord",
                    "WebView2"
                );

                Directory.CreateDirectory(userDataFolder);

                CoreWebView2Environment environment =
                    await CoreWebView2Environment.CreateAsync(
                        browserExecutableFolder: null,
                        userDataFolder: userDataFolder
                    );

                await webView.EnsureCoreWebView2Async(environment);

                ConfigureWebView();

                webView.Source = new Uri(
                    "https://giaotrinh-word.vercel.app/taskpane.html"
                );
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Không thể mở Giáo Trình Word.\n\n" +
                    ex.Message,
                    "Lỗi WebView2",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        /// <summary>
        /// Cấu hình trình duyệt nhúng.
        /// </summary>
        private void ConfigureWebView()
        {
            if (webView?.CoreWebView2 == null)
                return;

            // Cho phép JavaScript
            webView.CoreWebView2.Settings.IsScriptEnabled = true;

            // Không hiện menu chuột phải mặc định của Edge
            webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;

            // Không cho user zoom bằng Ctrl + bánh xe
            webView.CoreWebView2.Settings.IsZoomControlEnabled = false;

            // Không hiện thanh trạng thái URL
            webView.CoreWebView2.Settings.IsStatusBarEnabled = false;

            // Có thể bật lại lúc debug nếu cần
            webView.CoreWebView2.Settings.AreDevToolsEnabled = true;

            // Xử lý lỗi điều hướng
            webView.CoreWebView2.NavigationCompleted +=
                CoreWebView2_NavigationCompleted;

            // Chuẩn bị cho bridge Web -> C# sau này
            webView.CoreWebView2.WebMessageReceived +=
                CoreWebView2_WebMessageReceived;
        }

        private void CoreWebView2_NavigationCompleted(
            object sender,
            CoreWebView2NavigationCompletedEventArgs e)
        {
            if (!e.IsSuccess)
            {
                MessageBox.Show(
                    "Không tải được trang Giáo Trình Word.\n\n" +
                    $"WebErrorStatus: {e.WebErrorStatus}",
                    "Lỗi tải trang",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
            }
        }

        /// <summary>
        /// Nhận message từ JavaScript.
        /// Sau này dùng để điều khiển Word.
        /// </summary>
        private void CoreWebView2_WebMessageReceived(
            object sender,
            CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                string message = e.TryGetWebMessageAsString();

                if (string.IsNullOrWhiteSpace(message))
                    return;

                switch (message)
                {
                    case "insert-test-text":
                        Globals.ThisAddIn.Application.Selection.TypeText(
                            "Xin chào từ Giáo Trình Word!"
                        );
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    ex.Message,
                    "Lỗi xử lý lệnh",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        /// <summary>
        /// Reload trang web.
        /// </summary>
        public void ReloadPage()
        {
            if (webView?.CoreWebView2 != null)
            {
                webView.CoreWebView2.Reload();
            }
        }

        /// <summary>
        /// Điều hướng về trang chính của Giáo Trình.
        /// </summary>
        public void GoHome()
        {
            if (webView?.CoreWebView2 != null)
            {
                webView.Source = new Uri(
                    "https://giaotrinh-word.vercel.app/taskpane.html"
                );
            }
        }
    }
}
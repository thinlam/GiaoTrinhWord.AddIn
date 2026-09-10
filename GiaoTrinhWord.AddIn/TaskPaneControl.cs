using System;
using System.Collections.Generic;
using System.IO;
using System.Web.Script.Serialization;
using System.Windows.Forms;

using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

using Word = Microsoft.Office.Interop.Word;


namespace GiaoTrinhWord.AddIn
{
    public partial class TaskPaneControl : UserControl
    {
        private WebView2 webView;

        private readonly JavaScriptSerializer json =
            new JavaScriptSerializer();


        // Phạm vi của lần kiểm tra gần nhất.
        // Dùng để áp dụng gợi ý đúng trong vùng đã được kiểm tra.
        private int lastCheckStart = -1;
        private int lastCheckEnd = -1;
        private string lastCheckDocumentName = null;


        public TaskPaneControl()
        {
            InitializeComponent();

            InitializeWebViewControl();

            this.Load += TaskPaneControl_Load;
        }


        // =====================================================
        // INIT WEBVIEW CONTROL
        // =====================================================

        private void InitializeWebViewControl()
        {
            webView = new WebView2
            {
                Dock = DockStyle.Fill,
                Name = "webViewGiaoTrinh"
            };

            Controls.Add(webView);
        }


        // =====================================================
        // LOAD
        // =====================================================

        private async void TaskPaneControl_Load(
            object sender,
            EventArgs e
        )
        {
            try
            {
                string userDataFolder =
                    Path.Combine(
                        Environment.GetFolderPath(
                            Environment.SpecialFolder.LocalApplicationData
                        ),
                        "GiaoTrinhWord",
                        "WebView2"
                    );


                Directory.CreateDirectory(
                    userDataFolder
                );


                CoreWebView2Environment environment =
                    await CoreWebView2Environment.CreateAsync(
                        browserExecutableFolder: null,
                        userDataFolder: userDataFolder
                    );


                await webView.EnsureCoreWebView2Async(
                    environment
                );


                ConfigureWebView();


                // v=... giúp tránh cache JS cũ trong lúc phát triển
                webView.Source = new Uri(
                    "https://giaotrinh-word.vercel.app/taskpane.html?v=20260910-3"
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


        // =====================================================
        // CONFIG
        // =====================================================

        private void ConfigureWebView()
        {
            if (webView?.CoreWebView2 == null)
            {
                return;
            }


            webView.CoreWebView2.Settings.IsScriptEnabled =
                true;


            webView.CoreWebView2.Settings
                .AreDefaultContextMenusEnabled =
                false;


            webView.CoreWebView2.Settings
                .IsZoomControlEnabled =
                false;


            webView.CoreWebView2.Settings
                .IsStatusBarEnabled =
                false;


#if DEBUG
            webView.CoreWebView2.Settings
                .AreDevToolsEnabled =
                true;
#else
            webView.CoreWebView2.Settings
                .AreDevToolsEnabled =
                false;
#endif


            webView.CoreWebView2.NavigationCompleted +=
                CoreWebView2_NavigationCompleted;


            webView.CoreWebView2.WebMessageReceived +=
                CoreWebView2_WebMessageReceived;
        }


        // =====================================================
        // NAVIGATION
        // =====================================================

        private void CoreWebView2_NavigationCompleted(
            object sender,
            CoreWebView2NavigationCompletedEventArgs e
        )
        {
            if (e.IsSuccess)
            {
                return;
            }


            MessageBox.Show(
                "Không tải được trang Giáo Trình Word.\n\n" +
                $"WebErrorStatus: {e.WebErrorStatus}",
                "Lỗi tải trang",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );
        }


        // =====================================================
        // WEB -> VSTO
        // =====================================================

        private void CoreWebView2_WebMessageReceived(
            object sender,
            CoreWebView2WebMessageReceivedEventArgs e
        )
        {
            try
            {
                string rawJson =
                    e.WebMessageAsJson;


                if (string.IsNullOrWhiteSpace(rawJson))
                {
                    return;
                }


                Dictionary<string, object> request =
                    json.Deserialize<
                        Dictionary<string, object>
                    >(rawJson);


                if (
                    request == null ||
                    !request.ContainsKey("type")
                )
                {
                    return;
                }


                string type =
                    Convert.ToString(
                        request["type"]
                    );


                string requestId =
                    request.ContainsKey("requestId")
                        ? Convert.ToString(
                            request["requestId"]
                        )
                        : null;


                switch (type)
                {
                    // =========================================
                    // ĐỌC VĂN BẢN WORD
                    // =========================================

                    case "GET_DOCUMENT_TEXT":

                        HandleGetDocumentText(
                            requestId
                        );

                        break;


                    // =========================================
                    // CHẤP NHẬN GỢI Ý
                    // =========================================

                    case "ACCEPT_ISSUE":

                        HandleAcceptIssue(
                            requestId,
                            request
                        );

                        break;


                    // =========================================
                    // BỎ QUA GỢI Ý
                    // =========================================

                    case "IGNORE_ISSUE":

                        HandleIgnoreIssue(
                            requestId
                        );

                        break;


                    // =========================================
                    // ANNOTATIONS
                    // Hiện tại chưa highlight trực tiếp trong Word.
                    // Trả success để phía web không bị timeout.
                    // =========================================

                    case "APPLY_ANNOTATIONS":

                        SendSuccess(
                            requestId,
                            new
                            {
                                success = true
                            }
                        );

                        break;


                    // =========================================
                    // TEST BRIDGE
                    // =========================================

                    case "INSERT_TEST_TEXT":

                        Globals.ThisAddIn.Application
                            .Selection
                            .TypeText(
                                "Xin chào từ Giáo Trình Word!"
                            );


                        SendSuccess(
                            requestId,
                            new
                            {
                                success = true
                            }
                        );

                        break;


                    default:

                        SendError(
                            requestId,
                            "Lệnh không được hỗ trợ: " +
                            type
                        );

                        break;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    "WebMessageReceived Error: " +
                    ex
                );
            }
        }


        // =====================================================
        // GET DOCUMENT TEXT
        // =====================================================

        private void HandleGetDocumentText(
            string requestId
        )
        {
            try
            {
                Word.Application app =
                    Globals.ThisAddIn.Application;


                if (app.Documents.Count == 0)
                {
                    SendError(
                        requestId,
                        "Không có tài liệu Word đang mở."
                    );

                    return;
                }


                Word.Document document =
                    app.ActiveDocument;


                if (document == null)
                {
                    SendError(
                        requestId,
                        "Không tìm thấy tài liệu đang hoạt động."
                    );

                    return;
                }


                Word.Selection selection =
                    app.Selection;


                bool isSelection =
                    HasRealSelection(
                        selection
                    );


                Word.Range sourceRange =
                    isSelection
                        ? selection.Range.Duplicate
                        : document.Content.Duplicate;


                string text =
                    CleanWordText(
                        sourceRange.Text
                    );


                // Lưu phạm vi kiểm tra để nút "Chấp nhận"
                // ưu tiên sửa đúng vùng vừa được kiểm tra.
                lastCheckStart =
                    sourceRange.Start;

                lastCheckEnd =
                    sourceRange.End;

                lastCheckDocumentName =
                    document.Name;


                if (
                    string.IsNullOrWhiteSpace(
                        text
                    )
                )
                {
                    SendError(
                        requestId,
                        "Tài liệu chưa có nội dung để kiểm tra."
                    );

                    return;
                }


                SendSuccess(
                    requestId,
                    new
                    {
                        text = text,

                        isSelection =
                            isSelection,

                        documentName =
                            document.Name
                    }
                );
            }
            catch (Exception ex)
            {
                SendError(
                    requestId,
                    "Không thể đọc nội dung Word: " +
                    ex.Message
                );
            }
        }


        // =====================================================
        // ACCEPT ISSUE
        // =====================================================

        private void HandleAcceptIssue(
            string requestId,
            Dictionary<string, object> request
        )
        {
            try
            {
                Dictionary<string, object> data =
                    GetDictionary(
                        request,
                        "data"
                    );


                Dictionary<string, object> issue =
                    GetDictionary(
                        data,
                        "issue"
                    );


                string paragraphId =
                    GetString(
                        issue,
                        "paragraphId"
                    );


                string original =
                    GetString(
                        issue,
                        "original"
                    );


                string replacement =
                    GetString(
                        issue,
                        "replacement"
                    );


                if (
                    string.IsNullOrWhiteSpace(
                        original
                    )
                )
                {
                    SendError(
                        requestId,
                        "Không có nội dung gốc để thay thế."
                    );

                    return;
                }


                bool success =
                    ReplaceIssueInWord(
                        paragraphId,
                        original,
                        replacement ?? string.Empty
                    );


                SendSuccess(
                    requestId,
                    new
                    {
                        success = success
                    }
                );
            }
            catch (Exception ex)
            {
                SendError(
                    requestId,
                    "Không thể áp dụng gợi ý: " +
                    ex.Message
                );
            }
        }


        // =====================================================
        // IGNORE ISSUE
        // =====================================================

        private void HandleIgnoreIssue(
            string requestId
        )
        {
            SendSuccess(
                requestId,
                new
                {
                    success = true
                }
            );
        }


        // =====================================================
        // REPLACE ISSUE IN WORD
        // =====================================================

        private bool ReplaceIssueInWord(
            string paragraphId,
            string original,
            string replacement
        )
        {
            Word.Application app =
                Globals.ThisAddIn.Application;


            if (app.Documents.Count == 0)
            {
                return false;
            }


            Word.Document document =
                app.ActiveDocument;


            if (document == null)
            {
                return false;
            }


            // Không sửa nhầm nếu user đã chuyển sang tài liệu khác.
            if (
                !string.IsNullOrWhiteSpace(
                    lastCheckDocumentName
                ) &&
                !string.Equals(
                    document.Name,
                    lastCheckDocumentName,
                    StringComparison.OrdinalIgnoreCase
                )
            )
            {
                return false;
            }


            // 1. Ưu tiên đúng paragraph của lần kiểm tra.
            int paragraphIndex =
                ParseParagraphIndex(
                    paragraphId
                );


            if (
                paragraphIndex > 0 &&
                lastCheckStart >= 0 &&
                lastCheckEnd >= lastCheckStart
            )
            {
                Word.Range paragraphRange =
                    FindNonEmptyParagraphInLastCheckScope(
                        document,
                        paragraphIndex
                    );


                if (
                    paragraphRange != null &&
                    ReplaceInsideRange(
                        paragraphRange,
                        original,
                        replacement
                    )
                )
                {
                    UpdateLastCheckEndAfterReplacement(
                        original,
                        replacement
                    );

                    return true;
                }
            }


            // 2. Fallback: tìm trong toàn phạm vi vừa kiểm tra.
            if (
                lastCheckStart >= 0 &&
                lastCheckEnd >= lastCheckStart
            )
            {
                int documentEnd =
                    document.Content.End;


                int safeStart =
                    Math.Max(
                        0,
                        Math.Min(
                            lastCheckStart,
                            documentEnd
                        )
                    );


                int safeEnd =
                    Math.Max(
                        safeStart,
                        Math.Min(
                            lastCheckEnd,
                            documentEnd
                        )
                    );


                Word.Range checkedRange =
                    document.Range(
                        safeStart,
                        safeEnd
                    );


                if (
                    ReplaceInsideRange(
                        checkedRange,
                        original,
                        replacement
                    )
                )
                {
                    UpdateLastCheckEndAfterReplacement(
                        original,
                        replacement
                    );

                    return true;
                }
            }


            // 3. Fallback cuối: tìm trong toàn bộ document.
            Word.Range documentRange =
                document.Content.Duplicate;


            bool found =
                ReplaceInsideRange(
                    documentRange,
                    original,
                    replacement
                );


            if (found)
            {
                UpdateLastCheckEndAfterReplacement(
                    original,
                    replacement
                );
            }


            return found;
        }


        // =====================================================
        // REPLACE INSIDE RANGE
        // =====================================================

        private bool ReplaceInsideRange(
            Word.Range range,
            string original,
            string replacement
        )
        {
            if (
                range == null ||
                string.IsNullOrWhiteSpace(
                    original
                )
            )
            {
                return false;
            }


            Word.Range searchRange =
                range.Duplicate;


            Word.Find find =
                searchRange.Find;


            find.ClearFormatting();

            find.Text =
                original;

            find.Forward =
                true;

            find.Wrap =
                Word.WdFindWrap.wdFindStop;

            find.Format =
                false;

            find.MatchCase =
                false;

            find.MatchWholeWord =
                false;


            bool found =
                find.Execute();


            if (!found)
            {
                return false;
            }


            searchRange.Text =
                replacement;


            return true;
        }


        // =====================================================
        // FIND PARAGRAPH IN LAST CHECK SCOPE
        // paragraphId: vsto-paragraph-1, vsto-paragraph-2...
        // =====================================================

        private Word.Range FindNonEmptyParagraphInLastCheckScope(
            Word.Document document,
            int targetIndex
        )
        {
            if (
                document == null ||
                targetIndex <= 0 ||
                lastCheckStart < 0 ||
                lastCheckEnd < lastCheckStart
            )
            {
                return null;
            }


            int documentEnd =
                document.Content.End;


            int safeStart =
                Math.Max(
                    0,
                    Math.Min(
                        lastCheckStart,
                        documentEnd
                    )
                );


            int safeEnd =
                Math.Max(
                    safeStart,
                    Math.Min(
                        lastCheckEnd,
                        documentEnd
                    )
                );


            Word.Range scopeRange =
                document.Range(
                    safeStart,
                    safeEnd
                );


            int currentIndex =
                0;


            foreach (
                Word.Paragraph paragraph
                in scopeRange.Paragraphs
            )
            {
                int paragraphStart =
                    Math.Max(
                        paragraph.Range.Start,
                        safeStart
                    );


                int paragraphEnd =
                    Math.Min(
                        paragraph.Range.End,
                        safeEnd
                    );


                if (
                    paragraphEnd <
                    paragraphStart
                )
                {
                    continue;
                }


                Word.Range paragraphRange =
                    document.Range(
                        paragraphStart,
                        paragraphEnd
                    );


                string paragraphText =
                    CleanWordText(
                        paragraphRange.Text
                    );


                if (
                    string.IsNullOrWhiteSpace(
                        paragraphText
                    )
                )
                {
                    continue;
                }


                currentIndex++;


                if (
                    currentIndex ==
                    targetIndex
                )
                {
                    return paragraphRange;
                }
            }


            return null;
        }


        // =====================================================
        // PARSE PARAGRAPH ID
        // =====================================================

        private int ParseParagraphIndex(
            string paragraphId
        )
        {
            if (
                string.IsNullOrWhiteSpace(
                    paragraphId
                )
            )
            {
                return -1;
            }


            const string prefix =
                "vsto-paragraph-";


            if (
                !paragraphId.StartsWith(
                    prefix,
                    StringComparison.OrdinalIgnoreCase
                )
            )
            {
                return -1;
            }


            string indexText =
                paragraphId.Substring(
                    prefix.Length
                );


            int index;


            if (
                int.TryParse(
                    indexText,
                    out index
                )
            )
            {
                return index;
            }


            return -1;
        }


        // =====================================================
        // UPDATE CHECK RANGE AFTER REPLACE
        // =====================================================

        private void UpdateLastCheckEndAfterReplacement(
            string original,
            string replacement
        )
        {
            if (
                lastCheckEnd < 0
            )
            {
                return;
            }


            int originalLength =
                original?.Length ?? 0;


            int replacementLength =
                replacement?.Length ?? 0;


            lastCheckEnd +=
                replacementLength -
                originalLength;


            if (
                lastCheckEnd <
                lastCheckStart
            )
            {
                lastCheckEnd =
                    lastCheckStart;
            }
        }


        // =====================================================
        // REAL SELECTION
        // =====================================================

        private bool HasRealSelection(
            Word.Selection selection
        )
        {
            return (
                selection != null &&
                selection.Type !=
                    Word.WdSelectionType.wdSelectionIP &&
                !string.IsNullOrWhiteSpace(
                    selection.Text
                )
            );
        }


        // =====================================================
        // JSON HELPERS
        // =====================================================

        private Dictionary<string, object> GetDictionary(
            Dictionary<string, object> source,
            string key
        )
        {
            if (
                source == null ||
                !source.ContainsKey(key) ||
                source[key] == null
            )
            {
                return new Dictionary<string, object>();
            }


            Dictionary<string, object> value =
                source[key]
                    as Dictionary<string, object>;


            return (
                value ??
                new Dictionary<string, object>()
            );
        }


        private string GetString(
            Dictionary<string, object> source,
            string key
        )
        {
            if (
                source == null ||
                !source.ContainsKey(key) ||
                source[key] == null
            )
            {
                return null;
            }


            return Convert.ToString(
                source[key]
            );
        }


        // =====================================================
        // CLEAN WORD TEXT
        // =====================================================

        private string CleanWordText(
            string text
        )
        {
            if (
                string.IsNullOrEmpty(
                    text
                )
            )
            {
                return string.Empty;
            }


            return text
                .Replace("\a", "")
                .Replace("\v", "\n")
                .Trim();
        }


        // =====================================================
        // SEND SUCCESS
        // =====================================================

        private void SendSuccess(
            string requestId,
            object data
        )
        {
            if (
                webView?.CoreWebView2 == null
            )
            {
                return;
            }


            var response =
                new
                {
                    responseTo =
                        requestId,

                    ok =
                        true,

                    data =
                        data
                };


            string responseJson =
                json.Serialize(
                    response
                );


            webView.CoreWebView2
                .PostWebMessageAsJson(
                    responseJson
                );
        }


        // =====================================================
        // SEND ERROR
        // =====================================================

        private void SendError(
            string requestId,
            string message
        )
        {
            if (
                webView?.CoreWebView2 == null
            )
            {
                return;
            }


            var response =
                new
                {
                    responseTo =
                        requestId,

                    ok =
                        false,

                    message =
                        message
                };


            string responseJson =
                json.Serialize(
                    response
                );


            webView.CoreWebView2
                .PostWebMessageAsJson(
                    responseJson
                );
        }


        // =====================================================
        // RELOAD
        // =====================================================

        public void ReloadPage()
        {
            if (
                webView?.CoreWebView2 != null
            )
            {
                webView.CoreWebView2.Reload();
            }
        }


        // =====================================================
        // HOME
        // =====================================================

        public void GoHome()
        {
            if (
                webView?.CoreWebView2 != null
            )
            {
                webView.Source =
                    new Uri(
                        "https://giaotrinh-word.vercel.app/taskpane.html?v=20260910-3"
                    );
            }
        }
    }
}
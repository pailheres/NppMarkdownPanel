using Microsoft.Web.WebView2.Core;
using PanelCommon;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Windows.Forms;

namespace Webview2Viewer
{
    public class Webview2WebbrowserControl : IWebbrowserControl
    {
        const string virtualHostProtocol = "http://";
        const string virtualHostName = "markdownpanel-virtualhost";
        const string CONFIG_FOLDER_NAME = "MarkdownPanel";
        private Microsoft.Web.WebView2.WinForms.WebView2 webView;
        private int lastVerticalScroll = 0;
        private bool webViewInitialized = false;

        public Action<string> StatusTextChangedAction { get; set; }
        public Action RenderingDoneAction { get; set; }

        private string currentBody;
        private string currentStyle;

        private string currentDocumentPath;

        private CoreWebView2Environment environment = null;

        public Webview2WebbrowserControl()
        {
            webView = null;
        }

        public async void Initialize(int zoomLevel)
        {
            var cacheDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), CONFIG_FOLDER_NAME, "webview2");
            webView = new Microsoft.Web.WebView2.WinForms.WebView2();
            var opt = new CoreWebView2EnvironmentOptions();
            environment = await CoreWebView2Environment.CreateAsync(null, cacheDir, opt);
            await webView.EnsureCoreWebView2Async(environment);

            webView.AccessibleName = "webView";
            webView.Name = "webView";
            webView.ZoomFactor = ConvertToZoomFactor(zoomLevel);
            webView.Source = new Uri("about:blank", UriKind.Absolute);
            webView.Location = new Point(1, 27);
            webView.Size = new Size(800, 424);
            webView.Dock = DockStyle.Fill;
            webView.TabIndex = 0;
            webView.NavigationStarting += OnWebBrowser_NavigationStarting;
            webView.NavigationCompleted += WebView_NavigationCompleted;
            webView.ZoomFactor = ConvertToZoomFactor(zoomLevel);

            webViewInitialized = true;
        }

        public void AddToHost(Control host)
        {
            host.Controls.Add(webView);
        }

        private async void WebView_NavigationCompleted(object sender, CoreWebView2NavigationCompletedEventArgs e)
        {
            // After NavigateToString completes, (re)render Mermaid first, then MathJax, then restore scroll.
            try
            {
                await RetypesetAsync();
            }
            catch { }

            ExecuteWebviewAction(new Action(async () =>
            {
                try
                {
                    await webView.ExecuteScriptAsync("window.scrollBy(0, " + lastVerticalScroll + " )");
                }
                catch { }
                if (RenderingDoneAction != null) RenderingDoneAction();
            }));
        }

        public Bitmap MakeScreenshot()
        {
            return null;
        }

        public void PrepareContentUpdate(bool preserveVerticalScrollPosition)
        {
            if (!webViewInitialized) return;
            if (preserveVerticalScrollPosition)
            {
                ExecuteWebviewAction(new Action(async () =>
                {
                    var scrollPosition = await webView.ExecuteScriptAsync("window.pageYOffset");
                    lastVerticalScroll = int.Parse(scrollPosition.Split('.')[0]);
                }));
            }
            else
            {
                lastVerticalScroll = 0;
            }
        }

        const string scrollScript =
            "var element = document.getElementById('{0}');\n" +
            "var headerOffset = 10;\n" +
            "var elementPosition = element.getBoundingClientRect().top;\n" +
            "var offsetPosition = elementPosition + window.pageYOffset - headerOffset;\n" +
            "window.scrollTo({{top: offsetPosition}});";

        public void ScrollToElementWithLineNo(int lineNo)
        {
            if (lineNo <= 0) lineNo = 0;
            ExecuteWebviewAction(new Action(async () =>
            {
                await webView.ExecuteScriptAsync(string.Format(scrollScript, lineNo));
            }));
        }

        public void SetContent(string content, string body, string style, string currentDocumentPath)
        {
            if (!webViewInitialized) return;

            var currentPath = Path.GetDirectoryName(currentDocumentPath);
            var replaceFileMapping = "file:///" + currentPath.Replace('\\', '/');

            content = content.Replace(replaceFileMapping, virtualHostProtocol + virtualHostName);
            body = body.Replace(replaceFileMapping, virtualHostProtocol + virtualHostName);

            var fullReload = false;
            if (this.currentDocumentPath != currentDocumentPath)
            {
                ExecuteWebviewAction(new Action(() =>
                {
                    webView.CoreWebView2.SetVirtualHostNameToFolderMapping(virtualHostName, currentPath, CoreWebView2HostResourceAccessKind.Allow);
                }));
                this.currentDocumentPath = currentDocumentPath;
                fullReload = true;
            }

            // Ensure <base> so relative URLs resolve against the MD folder
            const string baseTag = "<base href=\"http://markdownpanel-virtualhost/\">";
            if (content.IndexOf("<base", StringComparison.OrdinalIgnoreCase) < 0)
            {
                int headIdx = content.IndexOf("<head", StringComparison.OrdinalIgnoreCase);
                if (headIdx >= 0)
                {
                    int headClose = content.IndexOf('>', headIdx);
                    if (headClose >= 0)
                        content = content.Substring(0, headClose + 1) + baseTag + content.Substring(headClose + 1);
                }
                else
                {
                    content = "<head>" + baseTag + "</head>" + content;
                }
            }

            if (!fullReload && currentBody != null && currentStyle != null)
            {
                bool didChange = false;

                if (currentBody != body)
                {
                    currentBody = body;
                    didChange = true;
                    ExecuteWebviewAction(new Action(async () =>
                    {
                        await webView.ExecuteScriptAsync(
                            "document.body.innerHTML = '" + HttpUtility.JavaScriptStringEncode(currentBody) + "'"
                        );
                        // Re-render Mermaid then MathJax after DOM change
                        try { await RetypesetAsync(); } catch { }
                    }));
                }
                if (currentStyle != style)
                {
                    currentStyle = style;
                    didChange = true;
                    ExecuteWebviewAction(new Action(async () =>
                    {
                        await webView.ExecuteScriptAsync(
                            "document.head.removeChild(document.head.lastElementChild);\n" +
                            "var style = document.createElement('style');\n" +
                            "style.type = 'text/css'; \n" +
                            "style.textContent = '" + HttpUtility.JavaScriptStringEncode(currentStyle) + "'; \n" +
                            "document.head.appendChild(style); \n"
                        );
                        // Style changes can affect layout; re-typeset for good measure
                        try { await RetypesetAsync(); } catch { }
                    }));
                }

                // If neither body nor style changed, do nothing.
                if (!didChange)
                {
                    // No-op
                }
            }
            else
            {
                currentBody = body;
                currentStyle = style;
                ExecuteWebviewAction(new Action(() =>
                {
                    webView.NavigateToString(content);
                }));
            }
        }

        public void SetZoomLevel(int zoomLevel)
        {
            double zoomFactor = ConvertToZoomFactor(zoomLevel);
            ExecuteWebviewAction(new Action(() =>
            {
                if (webView.ZoomFactor != zoomFactor)
                    webView.ZoomFactor = zoomFactor;

            }));
        }

        private double ConvertToZoomFactor(int zoomLevel)
        {
            double zoomFactor = Convert.ToDouble(zoomLevel) / 100;
            return zoomFactor;
        }

        void OnWebBrowser_NavigationStarting(object sender, CoreWebView2NavigationStartingEventArgs e)
        {
            if (e.Uri.ToString().StartsWith("about:blank"))
            {
                e.Cancel = true;
            }
            else if (!e.Uri.ToString().StartsWith("data:"))
            {
                e.Cancel = true;
                var p = new Process();
                var navUri = e.Uri.ToString();
                if (navUri.StartsWith(virtualHostProtocol + virtualHostName))
                {
                    var currentPath = Path.GetDirectoryName(currentDocumentPath);
                    navUri = navUri.Replace(virtualHostProtocol + virtualHostName, currentPath);
                    navUri = Uri.UnescapeDataString(navUri);
                }
                p.StartInfo = new ProcessStartInfo(navUri);
                p.Start();
            }
        }

        public string GetRenderingEngineName()
        {
            return "EDGE";
        }

        private void ExecuteWebviewAction(Action action)
        {
            try
            {
                webView.Invoke(action);
            }
            catch (Exception)
            {
            }
        }

        // --- Added: centralized re-render for Mermaid then MathJax ---
        private async Task RetypesetAsync()
        {
            if (webView == null || webView.CoreWebView2 == null) return;

            const string js = @"
(async function(){
  try {
    if (window.__renderMermaid) {
      await window.__renderMermaid();
    }
  } catch(e) { /* ignore */ }

  try {
    if (window.retypesetMath) {
      await window.retypesetMath();
    } else if (window.MathJax && MathJax.typesetPromise) {
      await MathJax.typesetPromise();
    }
  } catch(e) { /* ignore */ }
})();";
            await webView.ExecuteScriptAsync(js);
        }
    }
}

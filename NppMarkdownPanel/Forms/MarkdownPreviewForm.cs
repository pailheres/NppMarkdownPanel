using NppMarkdownPanel.Entities;
using NppMarkdownPanel.Generator;
using NppMarkdownPanel.Webbrowser;
using PanelCommon;
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Webview2Viewer;

namespace NppMarkdownPanel.Forms
{
    public partial class MarkdownPreviewForm : Form, IViewerInterface
    {
        const string DEFAULT_HTML_BASE =
        @"<!DOCTYPE html>
          <html>
            <head>
              <meta http-equiv=""X-UA-Compatible"" content=""IE=edge"">
              <meta http-equiv=""content-type"" content=""text/html; charset=utf-8"">
              <meta name=""color-scheme"" content=""light"">
              <title>{0}</title>
              <style type=""text/css"">
                /* keep UA in light mode and force white canvas */
                :root {{ color-scheme: light; }}
                html, body {{ background:#fff !important; color:#111 !important; }}
                .mermaid {{ background:#fff !important; color:#111 !important; }}
                {1}
              </style>
            </head>
            <body class=""markdown-body"" style=""{2}"">
            {3}
            </body>
          </html>
          ";

        // inject Mermaid only into the browser preview (and optionally export)
        const string MermaidScript =
            @"<script type=""module"">
              import mermaid from 'https://cdn.jsdelivr.net/npm/mermaid@11/dist/mermaid.esm.min.mjs';
                window.__renderMermaid = () => {
                  const nodes = document.querySelectorAll('pre code.language-mermaid, pre code.mermaid');
                  nodes.forEach(code => {
                    const graph = code.textContent;
                    const pre = code.closest('pre') || code;
                    const div = document.createElement('div');
                    div.className = 'mermaid';
                    div.textContent = graph;
                    pre.replaceWith(div);
                  });
                  mermaid.initialize({ startOnLoad: true, theme: 'default' });  // not 'dark'
                  mermaid.run({ querySelector: '.mermaid' });
                };
                document.addEventListener('DOMContentLoaded', () => { try { window.__renderMermaid(); } catch(e) {} });
              </script>";

        const string MathJaxScript = @"
<script>
  window.MathJax = {
    tex: {
      inlineMath: [['$','$'], ['\\(','\\)']],
      displayMath: [['$$','$$'], ['\\[','\\]']],
      processEscapes: true,
      tags: 'ams'
    },
    options: {
      skipHtmlTags: ['script','noscript','style','textarea','pre','code']
    },
    startup: { typeset: false }   // we'll call it manually when ready
  };
</script>
<script src=""https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-chtml.js"" defer></script>
<script>
  (function waitMJ(){
    // Wait for MathJax to finish bootstrapping, then typeset the document
    if (window.MathJax && MathJax.startup && MathJax.typesetPromise) {
      MathJax.startup.promise.then(function(){ MathJax.typesetPromise(); });
    } else {
      setTimeout(waitMJ, 50);
    }
  })();
</script>";


        const string MSG_NO_SUPPORTED_FILE_EXT = "<h3>The current file <u>{0}</u> has no valid Markdown file extension.</h3><div>Valid file extensions:{1}</div>";

        private Task<RenderResult> renderTask;

        private string htmlContentForExport;
        private Settings settings;
        private string currentFilePath;
        private IWebbrowserControl webbrowserControl;
        private IWebbrowserControl webview1Instance;
        private IWebbrowserControl webview2Instance;

        private static readonly Regex IncludeRe =
    new Regex(@"<!--\s*@include\s+""(?<path>[^""]+)""\s*(?<opts>[^>]*)-->", RegexOptions.IgnoreCase);

        private string ExpandIncludes(string text, string currentFilePath, int depth = 0)
        {
            if (depth > 12) return text; // prevent infinite recursion
            string baseDir = string.IsNullOrEmpty(currentFilePath)
                ? Directory.GetCurrentDirectory()
                : Path.GetDirectoryName(currentFilePath);

            return IncludeRe.Replace(text, m =>
            {
                var spec = m.Groups["path"].Value.Trim();
                var opts = m.Groups["opts"].Value;
                int levelOffset = 0;

                // parse optional level=N
                var lv = Regex.Match(opts ?? "", @"\blevel\s*=\s*(?<n>-?\d+)");
                if (lv.Success) int.TryParse(lv.Groups["n"].Value, out levelOffset);

                // split out #section-id
                string filePart = spec, section = null;
                int hash = spec.IndexOf('#');
                if (hash >= 0) { filePart = spec.Substring(0, hash); section = spec.Substring(hash + 1); }

                string full = Path.GetFullPath(Path.Combine(baseDir, filePart));
                if (!File.Exists(full)) return $"<!-- include not found: {spec} -->";

                string content = File.ReadAllText(full);

                // recursively expand inner includes
                content = ExpandIncludes(content, full, depth + 1);

                // if a section is requested, extract it (very simple: from heading with id/text match until next same/greater level)
                if (!string.IsNullOrEmpty(section))
                    content = ExtractSectionByIdOrTitle(content, section);

                // apply heading offset
                if (levelOffset != 0)
                    content = Regex.Replace(content, @"(?m)^(#{1,6})(\s)", m2 =>
                    {
                        int n = Math.Max(1, Math.Min(6, m2.Groups[1].Value.Length + levelOffset));
                        return new string('#', n) + m2.Groups[2].Value;
                    });

                return content;
            });
        }

        // (Optional) very simple section extractor: matches ATX headings with id anchors like {#id} or GitHub-style (#id)
        private string ExtractSectionByIdOrTitle(string md, string section)
        {
            // try explicit attribute: ### Title {#section}
            var re = new Regex(@"(?ms)^(?<h>#{1,6})\s+(?<t>.+?)\s*(?:\{#(?<id>[^}]+)\})?\s*$");
            var lines = md.Split('\n');
            var sb = new StringBuilder();
            bool copy = false;
            int startLevel = 0;

            for (int i = 0; i < lines.Length; i++)
            {
                var line = lines[i];
                var m = re.Match(line);
                if (m.Success)
                {
                    var lvl = m.Groups["h"].Value.Length;
                    var id = m.Groups["id"].Success ? m.Groups["id"].Value.Trim() : null;
                    var title = Regex.Replace(m.Groups["t"].Value, @"\s*\{#.*\}\s*$", "").Trim();

                    // normalize GitHub-style id for title
                    string ghId = Regex.Replace(title.ToLowerInvariant(), @"[^\w\- ]+", "")
                                       .Replace(' ', '-');

                    if (!copy && (string.Equals(section, id, StringComparison.OrdinalIgnoreCase) ||
                                  string.Equals(section, ghId, StringComparison.OrdinalIgnoreCase)))
                    {
                        copy = true; startLevel = lvl;
                    }
                    else if (copy && lvl <= startLevel)
                    {
                        break; // reached next peer or higher heading
                    }
                }
                if (copy) sb.AppendLine(line);
            }
            return copy ? sb.ToString() : $"<!-- section not found: {section} -->";
        }

        public void UpdateSettings(Settings settings)
        {
            this.settings = settings;

            var isDarkModeEnabled = settings.IsDarkModeEnabled;
            if (isDarkModeEnabled)
            {
                tbPreview.BackColor = Color.Black;
                btnSaveHtml.ForeColor = Color.White;
                statusStrip2.BackColor = Color.Black;
                toolStripStatusLabel1.ForeColor = Color.White;
            }
            else
            {
                tbPreview.BackColor = SystemColors.Control;
                btnSaveHtml.ForeColor = SystemColors.ControlText;
                statusStrip2.BackColor = SystemColors.Control;
                toolStripStatusLabel1.ForeColor = SystemColors.ControlText;
            }

            tbPreview.Visible = settings.ShowToolbar;
            statusStrip2.Visible = settings.ShowStatusbar;

            if (webbrowserControl != null)
            {
                if (webbrowserControl.GetRenderingEngineName() != settings.RenderingEngine)
                {
                    InitRenderingEngine(settings);
                }
            }
        }

        private MarkdownService markdownService;
        private ActionRef<Message> wndProcCallback;

        public static IViewerInterface InitViewer(Settings settings, ActionRef<Message> wndProcCallback)
        {
            return new MarkdownPreviewForm(settings, wndProcCallback);
        }

        private MarkdownPreviewForm(Settings settings, ActionRef<Message> wndProcCallback)
        {
            InitializeComponent();

            this.wndProcCallback = wndProcCallback;
            markdownService = new MarkdownService(new MarkdigWrapper.MarkdigWrapper());
            markdownService.PreProcessorCommandFilename = settings.PreProcessorCommandFilename;
            markdownService.PreProcessorArguments = settings.PreProcessorArguments;
            markdownService.PostProcessorCommandFilename = settings.PostProcessorCommandFilename;
            markdownService.PostProcessorArguments = settings.PostProcessorArguments;
            this.settings = settings;
            panel1.Visible = true;

            InitRenderingEngine(settings);
        }

        private void InitRenderingEngine(Settings settings)
        {
            panel1.Controls.Clear();
          

            if (settings.IsRenderingEngineIE11())
            {
                if (webview1Instance == null)
                {
                    webbrowserControl = new IE11WebbrowserControl();
                    webbrowserControl.Initialize(settings.ZoomLevel);
                    webview1Instance = webbrowserControl;
                }
                else
                {
                    webbrowserControl = webview1Instance;
                }
            }
            else if (settings.IsRenderingEngineEdge())
            {
                if (webview2Instance == null)
                {
                    webbrowserControl = new Webview2WebbrowserControl();
                    webbrowserControl.Initialize(settings.ZoomLevel);
                    webview2Instance = webbrowserControl;
                }
                else
                {
                    webbrowserControl = webview2Instance;
                }
            }

            webbrowserControl.AddToHost(panel1);
            webbrowserControl.RenderingDoneAction = () => { HideScreenshotAndShowBrowser(); };
            webbrowserControl.StatusTextChangedAction = (status) => { toolStripStatusLabel1.Text = status; };
        }

        private RenderResult RenderHtmlInternal(string currentText, string filepath)
        {
            var defaultBodyStyle = "";
            var markdownStyleContent = GetCssContent(filepath);

            if (!IsValidFileExtension(currentFilePath))
            {
                var invalidExtensionMessageBody = string.Format(MSG_NO_SUPPORTED_FILE_EXT, Path.GetFileName(filepath), settings.SupportedFileExt);
                var invalidExtensionMessage = string.Format(DEFAULT_HTML_BASE, Path.GetFileName(filepath), markdownStyleContent, defaultBodyStyle, invalidExtensionMessageBody);

                return new RenderResult(invalidExtensionMessage, invalidExtensionMessage, invalidExtensionMessageBody, markdownStyleContent);
            }

            // ⬇️ NEW: expand <!-- @include "file.md" --> before converting to HTML
            var expanded = ExpandIncludes(currentText, filepath);

            var resultForBrowser = markdownService.ConvertToHtml(expanded, filepath, true);
            var resultForExport = markdownService.ConvertToHtml(expanded, null, false);

            var markdownHtmlBrowser = string.Format(DEFAULT_HTML_BASE, Path.GetFileName(filepath), markdownStyleContent, defaultBodyStyle, resultForBrowser);
            var markdownHtmlFileExport = string.Format(DEFAULT_HTML_BASE, Path.GetFileName(filepath), markdownStyleContent, defaultBodyStyle, resultForExport);

            markdownHtmlBrowser = markdownHtmlBrowser.Replace("</body>", MathJaxScript + MermaidScript + "</body>");
            // if you also want it in exported HTML, uncomment the next line
            // markdownHtmlFileExport = markdownHtmlFileExport.Replace("</body>", MathJaxScript + MermaidScript + "</body>");

            return new RenderResult(markdownHtmlBrowser, markdownHtmlFileExport, resultForBrowser, markdownStyleContent);
        }

        private string GetCssContent(string filepath)
        {
            // Path of plugin directory
            var cssContent = "";

            var assemblyPath = Path.GetDirectoryName(Assembly.GetAssembly(GetType()).Location);

            var defaultCss = settings.IsDarkModeEnabled ? Settings.DefaultDarkModeCssFile : Settings.DefaultCssFile;
            var customCssFile = settings.IsDarkModeEnabled ? settings.CssDarkModeFileName : settings.CssFileName;
            if (File.Exists(customCssFile))
            {
                cssContent = File.ReadAllText(customCssFile);
            }
            else
            {
                cssContent = File.ReadAllText(assemblyPath + "\\" + defaultCss);
            }

            return cssContent;
        }

        public void RenderMarkdown(string currentText, string filepath, bool preserveVerticalScrollPosition = true)
        {
            if (renderTask == null || renderTask.IsCompleted)
            {
                MakeAndDisplayScreenShot();
                webbrowserControl.PrepareContentUpdate(preserveVerticalScrollPosition);

                var context = TaskScheduler.FromCurrentSynchronizationContext();
                renderTask = new Task<RenderResult>(() => RenderHtmlInternal(currentText, filepath));
                renderTask.ContinueWith((renderedText) =>
                {
                    webbrowserControl.SetContent(renderedText.Result.ResultForBrowser, renderedText.Result.ResultBody, renderedText.Result.ResultStyle, currentFilePath);
                    htmlContentForExport = renderedText.Result.ResultForExport;
                    if (!String.IsNullOrWhiteSpace(settings.HtmlFileName))
                    {
                        bool valid = Utils.ValidateFileSelection(settings.HtmlFileName, out string fullPath, out string error, "HTML Output");
                        if (valid)
                        {
                            settings.HtmlFileName = fullPath; // the validation was run against this path, so we want to make sure the state of the preview matches that
                            writeHtmlContentToFile(settings.HtmlFileName);
                        }
                    }
                    webbrowserControl.SetZoomLevel(settings.ZoomLevel);

                }, context);
                renderTask.Start();
            }
        }
        /// <summary>
        /// Makes and displays a screenshot of the current browser content to prevent it from flickering 
        /// while loading updated content
        /// </summary>
        private void MakeAndDisplayScreenShot()
        {
            Bitmap bm = webbrowserControl.MakeScreenshot();
            if (bm != null)
            {
                pictureBoxScreenshot.Image = bm;
                pictureBoxScreenshot.Visible = true;
            }

        }

        private void HideScreenshotAndShowBrowser()
        {
            if (pictureBoxScreenshot.Image != null)
            {
                pictureBoxScreenshot.Visible = false;
                pictureBoxScreenshot.Image = null;
            }
        }

        public void ScrollToElementWithLineNo(int lineNo)
        {
            webbrowserControl.ScrollToElementWithLineNo((int)lineNo);
        }

        protected override void WndProc(ref Message m)
        {
            wndProcCallback(ref m);

            //Continue the processing, as we only toggle
            base.WndProc(ref m);
        }

        private void btnSaveHtml_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "html files (*.html, *.htm)|*.html;*.htm|All files (*.*)|*.*";
                saveFileDialog.RestoreDirectory = true;
                saveFileDialog.InitialDirectory = Path.GetDirectoryName(currentFilePath);
                saveFileDialog.FileName = Path.GetFileNameWithoutExtension(currentFilePath);
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    writeHtmlContentToFile(saveFileDialog.FileName);
                }
            }
        }

        private void writeHtmlContentToFile(string filename)
        {
            if (!string.IsNullOrEmpty(filename))
            {
                File.WriteAllText(filename, htmlContentForExport);
            }
        }

        public bool IsValidFileExtension(string filename)
        {
            if (settings.AllowAllExtensions) return true;
            var currentExtension = Path.GetExtension(filename).ToLower();
            var matchExtensionList = false;
            try
            {
                matchExtensionList = settings.SupportedFileExt.Split(',').Any(ext => ext != null && currentExtension.Equals("." + ext.Trim().ToLower()));
            }
            catch (Exception)
            {
            }

            return matchExtensionList;
        }

        public void SetMarkdownFilePath(string filepath)
        {
            currentFilePath = filepath;
        }
    }
}

using NppMarkdownPanel.Entities;
using NppMarkdownPanel.Generator;
using NppMarkdownPanel.Webbrowser;
using PanelCommon;
using System;
using System.Collections.Generic;
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

        // Mermaid (light)
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
                mermaid.initialize({ startOnLoad: true, theme: 'default' });
                mermaid.run({ querySelector: '.mermaid' });
              };
              document.addEventListener('DOMContentLoaded', () => { try { window.__renderMermaid(); } catch(e) {} });
            </script>";

        // MathJax (robust loader; no nullable syntax)
        const string MathJaxScript = @"
<script>
  window.MathJax = {
    tex: {
      inlineMath: [['$', '$'], ['\\(', '\\)']],
      displayMath: [['$$', '$$'], ['\\[', '\\]']],
      processEscapes: true,
      processEnvironments: true,   // align/gather/cases
      tags: 'ams'                  // numbering + \eqref
    },
    options: {
      skipHtmlTags: ['script','noscript','style','textarea','pre','code']
    },
    startup: { typeset: false }     // we will typeset manually
  };
</script>
<script src=""https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-chtml.js"" defer></script>
<script>
  // Wait for MathJax to be ready once
  (function waitMJ(){
    if (window.MathJax && MathJax.startup && MathJax.typesetPromise) {
      MathJax.startup.promise.then(function(){ MathJax.typesetPromise(); });
    } else {
      setTimeout(waitMJ, 50);
    }
  })();

  // Helper the C# side can call after DOM mutations:
  window.retypesetMath = function(){
    if (window.MathJax && MathJax.typesetPromise) {
      return MathJax.typesetPromise();
    }
    return Promise.resolve();
  };
</script>";

        const string MSG_NO_SUPPORTED_FILE_EXT =
            "<h3>The current file <u>{0}</u> has no valid Markdown file extension.</h3><div>Valid file extensions:{1}</div>";

        private Task<RenderResult> renderTask;

        private string htmlContentForExport;
        private Settings settings;
        private string currentFilePath;
        private IWebbrowserControl webbrowserControl;
        private IWebbrowserControl webview1Instance;
        private IWebbrowserControl webview2Instance;

        // <!-- @include "file.md#id" --> or <!-- @include "file.md{#id .cls}" --> with opts after -->
        private static readonly Regex IncludeRe =
            new Regex(@"<!--\s*@include\s+""(?<spec>[^""]+)""\s*(?<opts>[^>]*)-->",
                      RegexOptions.IgnoreCase | RegexOptions.Compiled);

        // key=value options parser (e.g., level=2)
        private static Dictionary<string, string> ParseOpts(string s)
        {
            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (Match m in Regex.Matches(s ?? "", @"(?<k>\w+)\s*=\s*(?:""(?<v>[^""]*)""|(?<v>\S+))"))
                dict[m.Groups["k"].Value] = m.Groups["v"].Value;
            return dict;
        }

        private string ExpandIncludes(string text, string currentFilePath, int depth = 0)
        {
            if (depth > 20) return text; // recursion guard

            string baseDir = string.IsNullOrEmpty(currentFilePath)
                ? Directory.GetCurrentDirectory()
                : Path.GetDirectoryName(currentFilePath);

            return IncludeRe.Replace(text, m =>
            {
                var spec = m.Groups["spec"].Value.Trim();
                var opts = ParseOpts(m.Groups["opts"].Value);

                // heading offset
                int levelOffset = 0;
                int tmp;
                string levelStr;
                if (opts.TryGetValue("level", out levelStr) && int.TryParse(levelStr, out tmp))
                {
                    levelOffset = tmp;
                }

                // parse "file{#id .class}" or "file#id"
                string specFile = spec;
                string wantId = null;
                string wantClass = null;

                var brace = Regex.Match(spec, @"^(?<file>[^{}#]+)\{(?<attrs>[^}]*)\}\s*$");
                if (brace.Success)
                {
                    specFile = brace.Groups["file"].Value.Trim();
                    string attrs = brace.Groups["attrs"].Value.Trim();
                    var idm = Regex.Match(attrs, @"#(?<id>[\w\-]+)");
                    if (idm.Success) wantId = idm.Groups["id"].Value;
                    var cls = Regex.Match(attrs, @"\.(?<c>[\w\-]+)");
                    if (cls.Success) wantClass = cls.Groups["c"].Value;
                }
                else
                {
                    int hash = spec.IndexOf('#');
                    if (hash >= 0)
                    {
                        specFile = spec.Substring(0, hash);
                        wantId = spec.Substring(hash + 1);
                    }
                }

                bool isGlob = specFile.IndexOfAny(new[] { '*', '?' }) >= 0;
                var outputs = new List<string>();

                if (isGlob)
                {
                    string specDir = Path.GetDirectoryName(specFile) ?? "";
                    string pattern = Path.GetFileName(specFile);
                    string searchDir = Path.GetFullPath(Path.Combine(baseDir, specDir));
                    if (!Directory.Exists(searchDir))
                        return "<!-- include glob dir not found: " + specDir + " -->";
                    var files = Directory.GetFiles(searchDir, pattern, SearchOption.TopDirectoryOnly).ToList();
                    files.Sort(StringComparer.OrdinalIgnoreCase);
                    foreach (var full in files)
                        outputs.Add(ProcessOne(full, wantId, wantClass, levelOffset, depth));
                }
                else
                {
                    string full = Path.GetFullPath(Path.Combine(baseDir, specFile));
                    outputs.Add(ProcessOne(full, wantId, wantClass, levelOffset, depth));
                }

                return string.Join(Environment.NewLine, outputs.Where(s2 => !string.IsNullOrEmpty(s2)));
            });

            string ProcessOne(string fullPath, string sectionId, string requireClass, int level, int d)
            {
                if (!File.Exists(fullPath))
                    return "<!-- include not found: " + Path.GetFileName(fullPath) + " -->";

                string content = File.ReadAllText(fullPath);

                // expand nested includes relative to this file
                content = ExpandIncludes(content, fullPath, d + 1);

                // Extract by id/class if requested
                if (!string.IsNullOrEmpty(sectionId) || !string.IsNullOrEmpty(requireClass))
                {
                    var extracted = ExtractSectionByIdClassOrTitle(content, sectionId, requireClass);
                    if (extracted != null) content = extracted;
                    else return "<!-- section not found: " + (sectionId ?? ("." + requireClass)) + " in " + Path.GetFileName(fullPath) + " -->";
                }

                // Apply heading offset
                if (level != 0)
                {
                    content = Regex.Replace(content, @"(?m)^(#{1,6})(\s+)", m2 =>
                    {
                        int n = m2.Groups[1].Value.Length + level;
                        if (n < 1) n = 1; else if (n > 6) n = 6;
                        return new string('#', n) + m2.Groups[2].Value;
                    });
                }

                return content;
            }
        }

        // Extract section starting at a heading that matches id OR class OR GH slug
        private static string ExtractSectionByIdClassOrTitle(string md, string wantId, string wantClass)
        {
            var re = new Regex(@"^(?<h>#{1,6})\s+(?<title>.+?)(?:\s*\{(?<attrs>[^}]*)\})?\s*$",
                               RegexOptions.Multiline);

            Func<string, string> Slug = s =>
            {
                s = s.ToLowerInvariant();
                s = Regex.Replace(s, @"[^\w\- ]+", "");
                s = Regex.Replace(s, @"\s+", "-");
                return s.Trim('-');
            };

            Func<string, Tuple<string, HashSet<string>>> ParseAttrs = a =>
            {
                string id = null;
                var classes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                if (!string.IsNullOrWhiteSpace(a))
                {
                    foreach (Match m in Regex.Matches(a, @"([#.])([\w\-]+)"))
                    {
                        if (m.Groups[1].Value == "#") id = m.Groups[2].Value;
                        else classes.Add(m.Groups[2].Value);
                    }
                }
                return Tuple.Create(id, classes);
            };

            int start = -1, end = md.Length, startLevel = 0;

            foreach (Match m in re.Matches(md))
            {
                int lvl = m.Groups["h"].Value.Length;
                string titleRaw = m.Groups["title"].Value;
                string title = Regex.Replace(titleRaw, @"\s*\{[^}]*\}\s*$", "").Trim();
                var attrs = m.Groups["attrs"].Success ? m.Groups["attrs"].Value : null;

                var parsed = ParseAttrs(attrs);
                string id = parsed.Item1;
                HashSet<string> classes = parsed.Item2;

                string gh = Slug(title);

                bool matches = false;
                if (!string.IsNullOrEmpty(wantId))
                    matches = string.Equals(wantId, id, StringComparison.OrdinalIgnoreCase)
                           || string.Equals(wantId, gh, StringComparison.OrdinalIgnoreCase);
                if (!matches && !string.IsNullOrEmpty(wantClass))
                    matches = classes.Contains(wantClass);

                if (start < 0 && matches)
                {
                    start = m.Index;
                    startLevel = lvl;
                }
                else if (start >= 0 && lvl <= startLevel)
                {
                    end = m.Index; break;
                }
            }

            return start >= 0 ? md.Substring(start, end - start) : null;
        }

        // Legacy wrapper (kept if other code calls it)
        private string ExtractSectionByIdOrTitle(string md, string section)
        {
            var s = ExtractSectionByIdClassOrTitle(md, section, null);
            return s ?? "<!-- section not found: " + section + " -->";
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

            // expand <!-- @include ... -->
            var expanded = ExpandIncludes(currentText, filepath);

            var resultForBrowser = markdownService.ConvertToHtml(expanded, filepath, true);
            var resultForExport = markdownService.ConvertToHtml(expanded, null, false);

            var markdownHtmlBrowser = string.Format(DEFAULT_HTML_BASE, Path.GetFileName(filepath), markdownStyleContent, defaultBodyStyle, resultForBrowser);
            var markdownHtmlFileExport = string.Format(DEFAULT_HTML_BASE, Path.GetFileName(filepath), markdownStyleContent, defaultBodyStyle, resultForExport);

            // ---- Inject into <head>: base href + MathJax + Mermaid ----
            var baseTag = MakeBaseHref(filepath);
            var headInjections = baseTag + MathJaxScript + MermaidScript;

            markdownHtmlBrowser = InjectIntoHead(markdownHtmlBrowser, headInjections);
            // If you also want export HTML to include math/mermaid, uncomment:
            // markdownHtmlFileExport = InjectIntoHead(markdownHtmlFileExport, headInjections);

            return new RenderResult(markdownHtmlBrowser, markdownHtmlFileExport, resultForBrowser, markdownStyleContent);
        }

        private static string InjectIntoHead(string html, string toInsert)
        {
            int i = html.IndexOf("</head>", StringComparison.OrdinalIgnoreCase);
            if (i >= 0) return html.Substring(0, i) + toInsert + html.Substring(i);

            // If no <head> found, try to put it before body start
            int j = html.IndexOf("<body", StringComparison.OrdinalIgnoreCase);
            if (j >= 0) return "<head>" + toInsert + "</head>" + html;

            // Last fallback: prepend
            return "<head>" + toInsert + "</head>" + html;
        }

        private static string RobustAppendBeforeBodyEnd(string html, string injection)
        {
            int i = html.LastIndexOf("</body>", StringComparison.OrdinalIgnoreCase);
            if (i >= 0) return html.Insert(i, injection);
            return html + injection + "</body></html>";
        }

        private string GetCssContent(string filepath)
        {
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

        private static string MakeBaseHref(string mdPath)
        {
            try
            {
                if (string.IsNullOrEmpty(mdPath)) return string.Empty;
                if (!System.IO.File.Exists(mdPath)) return string.Empty;

                string dir = System.IO.Path.GetDirectoryName(mdPath);
                if (string.IsNullOrEmpty(dir)) return string.Empty;
                if (!System.IO.Directory.Exists(dir)) return string.Empty;

                // Ensure trailing separator so Uri becomes .../ (not .../doc)
                if (!dir.EndsWith(System.IO.Path.DirectorySeparatorChar.ToString()))
                {
                    dir += System.IO.Path.DirectorySeparatorChar;
                }

                // Convert to file:/// URL with proper escaping
                var uri = new System.Uri(dir);
                return "<base href=\"" + uri.AbsoluteUri + "\">";
            }
            catch
            {
                return string.Empty;
            }
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
                        string fullPath, error;
                        bool valid = Utils.ValidateFileSelection(settings.HtmlFileName, out fullPath, out error, "HTML Output");
                        if (valid)
                        {
                            settings.HtmlFileName = fullPath;
                            writeHtmlContentToFile(settings.HtmlFileName);
                        }
                    }
                    webbrowserControl.SetZoomLevel(settings.ZoomLevel);

                }, context);
                renderTask.Start();
            }
        }

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

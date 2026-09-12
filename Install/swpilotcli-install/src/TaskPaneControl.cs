using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

namespace SwpilotCLIAddin
{
    public class TaskPaneControl : UserControl
    {
        #region Win32 API

        [DllImport("user32.dll")]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll")]
        private static extern bool MoveWindow(IntPtr hWnd, int x, int y, int nWidth, int nHeight, bool bRepaint);

        [DllImport("user32.dll")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter,
            int X, int Y, int cx, int cy, uint uFlags);

        // ── Console input injection ───────────────────────────────────────────
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool AttachConsole(uint dwProcessId);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool FreeConsole();

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr GetStdHandle(int nStdHandle);

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern bool WriteConsoleInput(IntPtr hConsoleInput,
            [In] INPUT_RECORD[] lpBuffer, uint nLength, out uint lpNumberOfEventsWritten);

        private const int STD_INPUT_HANDLE = -10;

        [StructLayout(LayoutKind.Explicit, Size = 20)]
        private struct INPUT_RECORD
        {
            [FieldOffset(0)] public ushort EventType;
            [FieldOffset(4)] public KEY_EVENT_RECORD KeyEvent;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct KEY_EVENT_RECORD
        {
            [MarshalAs(UnmanagedType.Bool)]
            public bool bKeyDown;
            public ushort wRepeatCount;
            public ushort wVirtualKeyCode;
            public ushort wVirtualScanCode;
            public char   UnicodeChar;
            public uint   dwControlKeyState;
        }

        private const int GWL_STYLE        = -16;
        private const int GWL_EXSTYLE      = -20;
        private const int WS_POPUP         = unchecked((int)0x80000000);
        private const int WS_CAPTION       = 0x00C00000;
        private const int WS_THICKFRAME    = 0x00040000;
        private const int WS_MINIMIZE      = 0x20000000;
        private const int WS_MAXIMIZE      = 0x01000000;
        private const int WS_SYSMENU       = 0x00080000;
        private const int WS_CHILD         = 0x40000000;
        private const int WS_VISIBLE       = 0x10000000;
        private const int WS_EX_APPWINDOW  = 0x00040000;
        private const int WS_EX_TOOLWINDOW = 0x00000080;
        private const uint SWP_FRAMECHANGED = 0x0020;
        private const uint SWP_SHOWWINDOW   = 0x0040;

        #endregion

        // ── UI Controls ──────────────────────────────────────────────────────
        private Panel    topBar;
        private Label    lblPath;
        private Panel    toolBar;
        private ComboBox toolCombo;
        private Panel    historyPanel;
        private ComboBox historyCombo;
        private Button   btnRefreshHistory;
        private Panel    terminalPanel;
        private Label    lblHint;

        // ── Tool list ────────────────────────────────────────────────────────
        // This local build is dedicated to Codex-backed SOLIDWORKS automation.
        // Keep shell options for diagnostics, but do not offer or default to Claude.
        private readonly string[] ToolNames = { "codex", "PS", "cmd" };
        private string currentTool = "codex";

        // ── About info (bump on each release) ────────────────────────────────
        private const string AboutVersion = "0.1.1";
        private const string AboutGitHub  = "https://github.com/arthurle3210/SwpilotCLI";
        private const string AboutEmail   = "swpilot.arthurle3210@gmail.com";

        private static readonly System.Drawing.Color ColBorder = System.Drawing.Color.FromArgb(85, 85, 85);

        // ── State ────────────────────────────────────────────────────────────
        private Process psProcess;
        private IntPtr  psHwnd = IntPtr.Zero;
        private PSTerminalControl psTerminal;
        private string  workingDirectory;
        private string  terminalTitle;

        // ── History ──────────────────────────────────────────────────────────
        private string selectedSessionId               = null;
        private string selectedSessionWorkingDirectory = null;
        private bool   historyPopulating               = false;

        private struct SessionInfo
        {
            public string   Id;
            public string   Title;
            public string   Cwd;
            public DateTime Time;

            public override string ToString()
            {
                if (Id == null) return "  (New chat)";
                return string.Format("  {0:MM/dd HH:mm}  {1}", Time, Title);
            }
        }

        // ─────────────────────────────────────────────────────────────────────
        public TaskPaneControl()
        {
            InitializeUI();
            string installRoot = ResolveInstallRoot();
            if (!string.IsNullOrEmpty(installRoot))
                SetWorkingDirectory(installRoot);
        }

        private static string ResolveInstallRoot()
        {
            string dir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            while (!string.IsNullOrEmpty(dir))
            {
                if (File.Exists(Path.Combine(dir, "TOOLS.md")))
                    return dir;
                string parent = Path.GetDirectoryName(dir);
                if (parent == dir) break;
                dir = parent;
            }
            return null;
        }

        private void InitializeUI()
        {
            this.Dock      = DockStyle.Fill;
            this.BackColor = System.Drawing.Color.FromArgb(30, 30, 30);

            // ── Top bar: path + browse ────────────────────────────────────────
            topBar = new Panel
            {
                Dock      = DockStyle.Top,
                Height    = 36,
                BackColor = System.Drawing.Color.FromArgb(37, 37, 38),
                Padding   = new Padding(4, 4, 4, 4)
            };

            lblPath = new Label
            {
                Text      = "  Initializing...",
                Dock      = DockStyle.Fill,
                ForeColor = System.Drawing.Color.FromArgb(120, 120, 120),
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
                Font      = new System.Drawing.Font("Consolas", 8.5f)
            };

            topBar.Controls.Add(lblPath);

            // ── Tool bar: single ComboBox ─────────────────────────────────────
            toolBar = new Panel
            {
                Dock      = DockStyle.Top,
                Height    = 32,
                BackColor = System.Drawing.Color.FromArgb(30, 30, 30),
                Padding   = new Padding(4, 4, 4, 4)
            };

            toolCombo = new ComboBox
            {
                Dock          = DockStyle.Left,
                Width         = 120,
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor     = System.Drawing.Color.FromArgb(45, 45, 48),
                ForeColor     = System.Drawing.Color.White,
                Font          = new System.Drawing.Font("Consolas", 9f, System.Drawing.FontStyle.Bold),
                FlatStyle     = FlatStyle.Flat,
                Enabled       = true
            };
            foreach (string t in ToolNames)
                toolCombo.Items.Add(t);
            toolCombo.SelectedItem = currentTool;
            toolCombo.SelectedIndexChanged += ToolCombo_SelectedIndexChanged;
            toolBar.Controls.Add(toolCombo);

            var btnAbout = new Button
            {
                Text      = "≡",
                Dock      = DockStyle.Right,
                Width     = 26,
                FlatStyle = FlatStyle.Flat,
                ForeColor = System.Drawing.Color.White,
                BackColor = System.Drawing.Color.FromArgb(62, 62, 66),
                Font      = new System.Drawing.Font("Segoe UI", 10f, System.Drawing.FontStyle.Bold)
            };
            btnAbout.FlatAppearance.BorderSize = 0;
            btnAbout.Click += (s, e) =>
            {
                MessageBox.Show(
                    "SwpilotCLI" + Environment.NewLine +
                    "Version " + AboutVersion + Environment.NewLine +
                    "by Arthurle3210" + Environment.NewLine + Environment.NewLine +
                    "GitHub:" + Environment.NewLine +
                    AboutGitHub + Environment.NewLine + Environment.NewLine +
                    "Contact / Bug report:" + Environment.NewLine +
                    AboutEmail,
                    "About SwpilotCLI",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            };
            toolBar.Controls.Add(btnAbout);

            // ── History bar (visible when claude is selected) ─────────────────
            historyPanel = new Panel
            {
                Dock      = DockStyle.Top,
                Height    = 26,
                BackColor = System.Drawing.Color.FromArgb(37, 37, 38),
                Padding   = new Padding(4, 2, 4, 2),
                Visible   = false
            };

            btnRefreshHistory = new Button
            {
                Text      = "↺",
                Dock      = DockStyle.Right,
                Width     = 26,
                FlatStyle = FlatStyle.Flat,
                ForeColor = System.Drawing.Color.White,
                BackColor = System.Drawing.Color.FromArgb(62, 62, 66),
                Font      = new System.Drawing.Font("Segoe UI", 9f),
                Cursor    = Cursors.Hand
            };
            btnRefreshHistory.FlatAppearance.BorderColor = ColBorder;
            btnRefreshHistory.Click += (s, e) => RefreshHistory();

            historyCombo = new ComboBox
            {
                Dock          = DockStyle.Fill,
                BackColor     = System.Drawing.Color.FromArgb(30, 30, 30),
                ForeColor     = System.Drawing.Color.FromArgb(200, 200, 200),
                Font          = new System.Drawing.Font("Consolas", 8f),
                DropDownStyle = ComboBoxStyle.DropDownList,
                FlatStyle     = FlatStyle.Flat
            };
            historyCombo.SelectedIndexChanged += HistoryCombo_SelectedIndexChanged;

            historyPanel.Controls.Add(historyCombo);
            historyPanel.Controls.Add(btnRefreshHistory);

            // ── Terminal panel ───────────────────────────────────────────────
            terminalPanel = new Panel
            {
                Dock      = DockStyle.Fill,
                BackColor = System.Drawing.Color.FromArgb(12, 12, 12)
            };

            lblHint = new Label
            {
                Text      = "Starting terminal...",
                Dock      = DockStyle.Fill,
                ForeColor = System.Drawing.Color.FromArgb(80, 80, 80),
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                Font      = new System.Drawing.Font("微軟正黑體", 9f)
            };
            terminalPanel.Controls.Add(lblHint);

            // 加入順序：Fill 先加，Top 後加（後加的 Top 在上面）
            this.Controls.Add(terminalPanel);
            this.Controls.Add(historyPanel);
            this.Controls.Add(toolBar);
            this.Controls.Add(topBar);
        }

        // ────────────────────────────────────────────────────────────────────
        //  Tool selection
        // ────────────────────────────────────────────────────────────────────

        private void ToolCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (toolCombo.SelectedItem == null) return;
            string tool = toolCombo.SelectedItem.ToString();
            if (tool == currentTool) return;
            SelectTool(tool);
        }

        private void SelectTool(string tool)
        {
            if (tool == currentTool) return;

            // Shut down the previous tool BEFORE updating currentTool, so
            // GetExitCommand() sees the old tool and sends the right exit
            // sequence (e.g. don't send "/exit" to a cmd shell). StartTerminal
            // below calls GracefulShutdown again, but it's a no-op once the
            // process is gone.
            GracefulShutdown();

            currentTool = tool;

            // Sync ComboBox if called externally (currentTool already updated → no loop)
            if (toolCombo.SelectedItem?.ToString() != tool)
                toolCombo.SelectedItem = tool;

            historyPanel.Visible = (currentTool == "claude" || currentTool == "codex");
            if ((currentTool == "claude" || currentTool == "codex") && !string.IsNullOrEmpty(workingDirectory))
                RefreshHistory();

            if (!string.IsNullOrEmpty(workingDirectory))
                StartTerminal();
        }

        // ────────────────────────────────────────────────────────────────────
        //  Set working directory
        // ────────────────────────────────────────────────────────────────────

        public void SetWorkingDirectory(string path)
        {
            workingDirectory = path;

            if (!Directory.Exists(workingDirectory))
                Directory.CreateDirectory(workingDirectory);

            lblPath.Text      = "  " + workingDirectory;
            lblPath.ForeColor = System.Drawing.Color.FromArgb(200, 200, 200);

            toolCombo.Enabled = true;

            if (currentTool == "claude" || currentTool == "codex")
            {
                historyPanel.Visible = true;
                RefreshHistory();
            }

            StartTerminal();
        }

        // ────────────────────────────────────────────────────────────────────
        //  History
        // ────────────────────────────────────────────────────────────────────

        private void RefreshHistory()
        {
            historyPopulating = true;
            historyCombo.Items.Clear();
            selectedSessionId = null;
            selectedSessionWorkingDirectory = null;
            string normalizedWorkingDir = NormalizePathForCompare(workingDirectory);

            historyCombo.Items.Add(new SessionInfo
            {
                Id    = null,
                Cwd   = null,
                Title = "(New chat)",
                Time  = DateTime.MinValue
            });

            if (!string.IsNullOrEmpty(workingDirectory))
            {
                try
                {
                    if (currentTool == "codex")
                        PopulateCodexHistory(normalizedWorkingDir);
                    else
                        PopulateClaudeHistory(normalizedWorkingDir);
                }
                catch { }
            }

            historyCombo.SelectedIndex = 0;
            historyPopulating = false;
        }

        private void PopulateClaudeHistory(string normalizedWorkingDir)
        {
            string claudeDir = Path.Combine(
                GetClaudeDataDirectory(),
                "projects",
                PathToClaudeHash(workingDirectory));

            if (!Directory.Exists(claudeDir))
                return;

            var files = new DirectoryInfo(claudeDir)
                .GetFiles("*.jsonl")
                .OrderByDescending(f => f.LastWriteTime)
                .Take(20);

            foreach (var file in files)
            {
                string sessionId = Path.GetFileNameWithoutExtension(file.Name);
                string sessionWorkingDirectory;
                if (!TryReadSessionWorkingDirectory(file.FullName, out sessionWorkingDirectory))
                    continue;
                if (!string.Equals(sessionWorkingDirectory, normalizedWorkingDir, StringComparison.OrdinalIgnoreCase))
                    continue;

                string title = ReadSessionCustomTitle(file.FullName);
                if (string.IsNullOrEmpty(title))
                    title = ReadSessionTitle(file.FullName);
                if (string.IsNullOrEmpty(title))
                    continue;

                historyCombo.Items.Add(new SessionInfo
                {
                    Id    = sessionId,
                    Title = title,
                    Cwd   = sessionWorkingDirectory,
                    Time  = file.LastWriteTime
                });
            }
        }

        private void PopulateCodexHistory(string normalizedWorkingDir)
        {
            string sessionsRoot = Path.Combine(GetCodexDataDirectory(), "sessions");
            if (!Directory.Exists(sessionsRoot))
                return;

            Dictionary<string, string> titleMap = LoadCodexHistoryTitleMap();
            var sessions = new List<SessionInfo>();

            foreach (string filePath in Directory.EnumerateFiles(sessionsRoot, "*.jsonl", SearchOption.AllDirectories))
            {
                string sessionId;
                string sessionWorkingDirectory;
                if (!TryReadCodexSessionMeta(filePath, out sessionId, out sessionWorkingDirectory))
                    continue;
                if (!string.Equals(sessionWorkingDirectory, normalizedWorkingDir, StringComparison.OrdinalIgnoreCase))
                    continue;

                string title;
                if (!titleMap.TryGetValue(sessionId, out title) || !IsUsableHistoryTitle(title))
                    title = sessionId;

                var file = new FileInfo(filePath);
                sessions.Add(new SessionInfo
                {
                    Id    = sessionId,
                    Title = title,
                    Cwd   = sessionWorkingDirectory,
                    Time  = file.LastWriteTime
                });
            }

            foreach (SessionInfo session in sessions
                .OrderByDescending(s => s.Time)
                .Take(20))
            {
                historyCombo.Items.Add(session);
            }
        }

        private static Dictionary<string, string> LoadHistoryDisplayMap()
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                string historyPath = Path.Combine(
                    GetClaudeDataDirectory(),
                    "history.jsonl");
                if (!File.Exists(historyPath))
                    return map;

                foreach (string line in File.ReadLines(historyPath))
                {
                    var idMatch = Regex.Match(line, "\"sessionId\"\\s*:\\s*\"([^\"]+)\"");
                    var textMatch = Regex.Match(line, "\"display\"\\s*:\\s*\"((?:[^\"\\\\]|\\\\.)*)\"");
                    if (!idMatch.Success || !textMatch.Success)
                        continue;

                    string sessionId = idMatch.Groups[1].Value;
                    string title = DecodeJsonText(textMatch.Groups[1].Value);
                    if (!IsUsableHistoryTitle(title))
                        continue;

                    if (title.Length > 50)
                        title = title.Substring(0, 50) + "...";

                    // Newer entries overwrite older entries.
                    map[sessionId] = title;
                }
            }
            catch { }

            return map;
        }

        private static Dictionary<string, string> LoadCodexHistoryTitleMap()
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                string historyPath = Path.Combine(
                    GetCodexDataDirectory(),
                    "history.jsonl");
                if (!File.Exists(historyPath))
                    return map;

                foreach (string line in File.ReadLines(historyPath))
                {
                    var idMatch = Regex.Match(line, "\"session_id\"\\s*:\\s*\"([^\"]+)\"");
                    if (!idMatch.Success)
                        continue;

                    var textMatch = Regex.Match(line, "\"text\"\\s*:\\s*\"((?:[^\"\\\\]|\\\\.)*)\"");
                    if (!textMatch.Success)
                        continue;

                    string title = DecodeJsonText(textMatch.Groups[1].Value);
                    if (!IsUsableHistoryTitle(title))
                        continue;

                    if (title.Length > 50)
                        title = title.Substring(0, 50) + "...";

                    // Keep latest title for the same session id.
                    map[idMatch.Groups[1].Value] = title;
                }
            }
            catch { }

            return map;
        }

        private static bool TryReadSessionWorkingDirectory(string sessionFilePath, out string normalizedWorkingDirectory)
        {
            normalizedWorkingDirectory = null;

            try
            {
                using (var reader = new StreamReader(sessionFilePath))
                {
                    string line;
                    int maxLines = 80;
                    while ((line = reader.ReadLine()) != null && maxLines-- > 0)
                    {
                        var cwdMatch = Regex.Match(line, "\"cwd\"\\s*:\\s*\"((?:[^\"\\\\]|\\\\.)*)\"");
                        if (!cwdMatch.Success)
                            continue;

                        string cwd = NormalizePathForCompare(DecodeJsonText(cwdMatch.Groups[1].Value));
                        if (string.IsNullOrEmpty(cwd))
                            continue;

                        normalizedWorkingDirectory = cwd;
                        return true;
                    }
                }
            }
            catch { }

            return TryReadSessionWorkingDirectoryFromHistory(
                Path.GetFileNameWithoutExtension(sessionFilePath),
                out normalizedWorkingDirectory);
        }

        private static bool TryResolveSessionWorkingDirectory(string sessionId, out string normalizedWorkingDirectory)
        {
            normalizedWorkingDirectory = null;
            if (string.IsNullOrWhiteSpace(sessionId))
                return false;

            try
            {
                string projectsRoot = Path.Combine(GetClaudeDataDirectory(), "projects");
                if (Directory.Exists(projectsRoot))
                {
                    foreach (string filePath in Directory.EnumerateFiles(
                        projectsRoot, sessionId + ".jsonl", SearchOption.AllDirectories))
                    {
                        if (TryReadSessionWorkingDirectory(filePath, out normalizedWorkingDirectory))
                            return true;
                    }
                }
            }
            catch { }

            return TryReadSessionWorkingDirectoryFromHistory(sessionId, out normalizedWorkingDirectory);
        }

        private static bool TryResolveCodexSessionWorkingDirectory(string sessionId, out string normalizedWorkingDirectory)
        {
            normalizedWorkingDirectory = null;
            if (string.IsNullOrWhiteSpace(sessionId))
                return false;

            try
            {
                string sessionsRoot = Path.Combine(GetCodexDataDirectory(), "sessions");
                if (!Directory.Exists(sessionsRoot))
                    return false;

                foreach (string filePath in Directory.EnumerateFiles(sessionsRoot, "*.jsonl", SearchOption.AllDirectories))
                {
                    string foundSessionId;
                    if (!TryReadCodexSessionMeta(filePath, out foundSessionId, out normalizedWorkingDirectory))
                        continue;
                    if (string.Equals(foundSessionId, sessionId, StringComparison.OrdinalIgnoreCase))
                        return !string.IsNullOrEmpty(normalizedWorkingDirectory);
                }
            }
            catch { }

            normalizedWorkingDirectory = null;
            return false;
        }

        private static bool TryReadCodexSessionMeta(
            string sessionFilePath,
            out string sessionId,
            out string normalizedWorkingDirectory)
        {
            sessionId = null;
            normalizedWorkingDirectory = null;

            try
            {
                using (var reader = new StreamReader(sessionFilePath))
                {
                    string line;
                    int maxLines = 80;
                    while ((line = reader.ReadLine()) != null && maxLines-- > 0)
                    {
                        if (!line.Contains("\"session_meta\""))
                            continue;

                        var idMatch = Regex.Match(line, "\"id\"\\s*:\\s*\"([^\"]+)\"");
                        var cwdMatch = Regex.Match(line, "\"cwd\"\\s*:\\s*\"((?:[^\"\\\\]|\\\\.)*)\"");
                        if (!idMatch.Success || !cwdMatch.Success)
                            continue;

                        sessionId = idMatch.Groups[1].Value;
                        normalizedWorkingDirectory = NormalizePathForCompare(
                            DecodeJsonText(cwdMatch.Groups[1].Value));

                        if (!string.IsNullOrWhiteSpace(sessionId) &&
                            !string.IsNullOrEmpty(normalizedWorkingDirectory))
                            return true;
                    }
                }
            }
            catch { }

            sessionId = null;
            normalizedWorkingDirectory = null;
            return false;
        }

        private static bool TryReadSessionWorkingDirectoryFromHistory(string sessionId, out string normalizedWorkingDirectory)
        {
            normalizedWorkingDirectory = null;
            if (string.IsNullOrWhiteSpace(sessionId))
                return false;

            try
            {
                string historyPath = Path.Combine(GetClaudeDataDirectory(), "history.jsonl");
                if (!File.Exists(historyPath))
                    return false;

                foreach (string line in File.ReadLines(historyPath))
                {
                    var idMatch = Regex.Match(line, "\"sessionId\"\\s*:\\s*\"([^\"]+)\"");
                    if (!idMatch.Success ||
                        !string.Equals(idMatch.Groups[1].Value, sessionId, StringComparison.OrdinalIgnoreCase))
                        continue;

                    var projectMatch = Regex.Match(line, "\"project\"\\s*:\\s*\"((?:[^\"\\\\]|\\\\.)*)\"");
                    if (!projectMatch.Success)
                        continue;

                    string project = NormalizePathForCompare(DecodeJsonText(projectMatch.Groups[1].Value));
                    if (string.IsNullOrEmpty(project))
                        continue;

                    // Keep the latest matched project path from history lines.
                    normalizedWorkingDirectory = project;
                }
            }
            catch { }

            return !string.IsNullOrEmpty(normalizedWorkingDirectory);
        }

        private static bool IsUsableHistoryTitle(string title)
        {
            if (string.IsNullOrWhiteSpace(title))
                return false;

            string t = title.Trim();
            if (t.Length == 0)
                return false;
            if (t.StartsWith("<"))
                return false;
            if (t.StartsWith("/"))
                return false;
            if (string.Equals(t, "exit", StringComparison.OrdinalIgnoreCase))
                return false;

            return true;
        }

        private static string DecodeJsonText(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;

            string decoded = text
                .Replace("\\\\", "\\")
                .Replace("\\\"", "\"")
                .Replace("\\n", " ")
                .Replace("\\r", "")
                .Replace("\\t", " ");

            decoded = Regex.Replace(decoded, @"\\u([0-9a-fA-F]{4})",
                m => ((char)Convert.ToInt32(m.Groups[1].Value, 16)).ToString());

            return decoded.Trim();
        }

        private static string NormalizePathForCompare(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return string.Empty;

            string normalized = path.Trim().Replace('/', '\\');
            try
            {
                normalized = Path.GetFullPath(normalized);
            }
            catch { }

            if (normalized.Length > 3)
                normalized = normalized.TrimEnd('\\');

            return normalized;
        }

        private static string GetClaudeDataDirectory()
        {
            string envPath = Environment.GetEnvironmentVariable("CLAUDE_CONFIG_DIR");
            if (!string.IsNullOrWhiteSpace(envPath))
            {
                string normalizedEnv = NormalizePathForCompare(envPath);
                if (!string.IsNullOrEmpty(normalizedEnv))
                    return normalizedEnv;
            }

            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                ".claude");
        }

        private static string GetCodexDataDirectory()
        {
            string envPath = Environment.GetEnvironmentVariable("CODEX_HOME");
            if (!string.IsNullOrWhiteSpace(envPath))
            {
                string normalizedEnv = NormalizePathForCompare(envPath);
                if (!string.IsNullOrEmpty(normalizedEnv))
                    return normalizedEnv;
            }

            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                ".codex");
        }

        private static string PathToClaudeHash(string path)
        {
            return Regex.Replace(path, @"[^a-zA-Z0-9]", "-");
        }

        private static string ReadSessionCustomTitle(string filePath)
        {
            try
            {
                using (var reader = new StreamReader(filePath))
                {
                    string line;
                    string latestTitle = null;
                    while ((line = reader.ReadLine()) != null)
                    {
                        if (!line.Contains("\"type\":\"custom-title\"") &&
                            !line.Contains("\"type\": \"custom-title\""))
                            continue;

                        var match = Regex.Match(line, "\"customTitle\"\\s*:\\s*\"((?:[^\"\\\\]|\\\\.)*)\"");
                        if (!match.Success)
                            continue;

                        string title = DecodeJsonText(match.Groups[1].Value);
                        if (!IsUsableHistoryTitle(title))
                            continue;

                        if (title.Length > 50)
                            title = title.Substring(0, 50) + "...";

                        // Keep the latest rename record in this session file.
                        latestTitle = title;
                    }

                    return latestTitle;
                }
            }
            catch { }

            return null;
        }

        private static string ReadSessionTitle(string filePath)
        {
            try
            {
                using (var reader = new StreamReader(filePath))
                {
                    string line;
                    int maxLines = 50;
                    while ((line = reader.ReadLine()) != null && maxLines-- > 0)
                    {
                        if (!line.Contains("\"type\":\"user\"") && !line.Contains("\"type\": \"user\""))
                            continue;
                        if (line.Contains("\"isMeta\":true"))
                            continue;

                        // Try array format: "text":"..."
                        var match = Regex.Match(line, "\"text\"\\s*:\\s*\"((?:[^\"\\\\]|\\\\.)*)\"");
                        // Try string format: "content":"..."
                        if (!match.Success)
                            match = Regex.Match(line, "\"content\"\\s*:\\s*\"((?:[^\"\\\\]|\\\\.)*)\"");
                        if (!match.Success) continue;

                        string text = match.Groups[1].Value
                            .Replace("\\n", " ").Replace("\\r", "").Replace("\\t", " ");

                        text = Regex.Replace(text, @"\\u([0-9a-fA-F]{4})",
                            m => ((char)Convert.ToInt32(m.Groups[1].Value, 16)).ToString());

                        text = text.Trim();
                        if (text.Length == 0 || text.StartsWith("<")) continue;
                        if (text.Length > 50) text = text.Substring(0, 50) + "…";
                        return text;
                    }
                }
            }
            catch { }
            return null;
        }

        private void HistoryCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (historyPopulating || historyCombo.SelectedItem == null) return;

            var info = (SessionInfo)historyCombo.SelectedItem;
            selectedSessionId = info.Id;
            selectedSessionWorkingDirectory = info.Cwd;

            if (!string.IsNullOrEmpty(workingDirectory))
                StartTerminal();
        }

        // ────────────────────────────────────────────────────────────────────
        //  Graceful shutdown — sends /exit or /quit before killing
        // ────────────────────────────────────────────────────────────────────

        private string GetExitCommand()
        {
            switch (currentTool)
            {
                case "claude":   return "/exit";
                case "codex":    return "/quit";
                default:         return null; // PS/cmd - kill the shell if needed
            }
        }

        // Inject a line into the PowerShell console's input buffer via WriteConsoleInput.
        // Works regardless of window focus since it writes directly to the kernel buffer.
        private void SendLineToConsole(string text)
        {
            if (psProcess == null || psProcess.HasExited) return;
            try
            {
                // SolidWorks is a GUI app with no console → FreeConsole is a no-op but safe
                FreeConsole();
                if (!AttachConsole((uint)psProcess.Id)) return;

                IntPtr hStdin = GetStdHandle(STD_INPUT_HANDLE);
                if (hStdin == IntPtr.Zero || hStdin == new IntPtr(-1)) { FreeConsole(); return; }

                string line = text + "\r";
                var records = new INPUT_RECORD[line.Length * 2];
                for (int i = 0; i < line.Length; i++)
                {
                    records[i * 2] = new INPUT_RECORD
                    {
                        EventType = 1, // KEY_EVENT
                        KeyEvent  = new KEY_EVENT_RECORD
                        {
                            bKeyDown     = true,
                            wRepeatCount = 1,
                            UnicodeChar  = line[i]
                        }
                    };
                    records[i * 2 + 1] = new INPUT_RECORD
                    {
                        EventType = 1,
                        KeyEvent  = new KEY_EVENT_RECORD
                        {
                            bKeyDown     = false,
                            wRepeatCount = 1,
                            UnicodeChar  = line[i]
                        }
                    };
                }

                WriteConsoleInput(hStdin, records, (uint)records.Length, out _);
                FreeConsole();
            }
            catch { }
        }

        private void EnsurePSTerminal()
        {
            if (psTerminal != null)
                return;

            psTerminal = new PSTerminalControl
            {
                Dock = DockStyle.Fill,
                Visible = false
            };
            terminalPanel.Controls.Add(psTerminal);
            psTerminal.BringToFront();
        }

        private void GracefulShutdown()
        {
            if (psTerminal != null)
            {
                psTerminal.StopSession();
                psTerminal.Visible = false;
            }

            if (psProcess == null || psProcess.HasExited)
            {
                if (psProcess != null) { psProcess.Dispose(); psProcess = null; }
                psHwnd = IntPtr.Zero;
                return;
            }

            string exitCmd = GetExitCommand();
            if (exitCmd != null)
            {
                SendLineToConsole(exitCmd);
                psProcess.WaitForExit(4000); // wait up to 4 s for graceful exit
            }

            if (!psProcess.HasExited)
            {
                // .NET Framework 4.8 has no Process.Kill(true), so use taskkill
                // to kill the whole process tree — otherwise nested shells
                // (e.g. cmd inside PowerShell) survive as orphaned windows.
                try
                {
                    var tk = Process.Start(new ProcessStartInfo("taskkill", "/T /F /PID " + psProcess.Id)
                    {
                        CreateNoWindow = true,
                        UseShellExecute = false
                    });
                    tk?.WaitForExit(2000);
                }
                catch { }
                if (!psProcess.HasExited)
                {
                    try { psProcess.Kill(); } catch { }
                }
                Thread.Sleep(200);
            }

            psProcess.Dispose();
            psProcess = null;
            psHwnd    = IntPtr.Zero;
        }

        // ────────────────────────────────────────────────────────────────────
        //  Terminal: start / embed
        // ────────────────────────────────────────────────────────────────────

        private void StartTerminal()
        {
            GracefulShutdown();

            lblHint.Visible = false;

            if (currentTool == "PS")
            {
                EnsurePSTerminal();
                psTerminal.Visible = true;
                psTerminal.BringToFront();
                psTerminal.StartSession(
                    workingDirectory,
                    null,
                    "exit",
                    "Starting PowerShell terminal...",
                    BuildPsLaunchCommandLine());
                return;
            }

            if (psTerminal != null)
                psTerminal.Visible = false;

            terminalTitle = "SwpilotCLI_" + Guid.NewGuid().ToString("N").Substring(0, 8);

            string safeDir   = workingDirectory.Replace("'", "''");
            string safeTitle = terminalTitle;
            string command   = BuildCommand(safeDir);

            var psi = new ProcessStartInfo
            {
                FileName        = "powershell.exe",
                Arguments       = string.Format(
                    "-NoExit -Command \"$host.UI.RawUI.WindowTitle = '{0}'; {1}\"",
                    safeTitle, command),
                UseShellExecute = false,
                CreateNoWindow  = false
            };

            psProcess = new Process { StartInfo = psi };
            psProcess.Start();

            var thread = new Thread(() =>
            {
                IntPtr hwnd = IntPtr.Zero;
                for (int i = 0; i < 100 && hwnd == IntPtr.Zero; i++)
                {
                    Thread.Sleep(100);
                    hwnd = FindWindow(null, safeTitle);
                }
                if (hwnd != IntPtr.Zero)
                    try { this.Invoke(new Action(() => EmbedWindow(hwnd))); }
                    catch { }
            });
            thread.IsBackground = true;
            thread.Start();
        }

        private string BuildCommand(string safeDir)
        {
            switch (currentTool)
            {
                case "claude":
                    string launchDir = safeDir;
                    string resumeArg = selectedSessionId != null
                        ? " --resume " + selectedSessionId
                        : "";

                    if (!string.IsNullOrEmpty(selectedSessionId))
                    {
                        string resolvedSessionDir = selectedSessionWorkingDirectory;
                        if (string.IsNullOrEmpty(resolvedSessionDir))
                            TryResolveSessionWorkingDirectory(selectedSessionId, out resolvedSessionDir);

                        if (!string.IsNullOrEmpty(resolvedSessionDir))
                            launchDir = resolvedSessionDir.Replace("'", "''");
                    }

                    return string.Format("Set-Location -LiteralPath '{0}'; claude{1}", launchDir, resumeArg);
                case "codex":
                    string codexLaunchDir = safeDir;
                    string codexResumeArg = "";

                    if (!string.IsNullOrEmpty(selectedSessionId))
                    {
                        codexResumeArg = " resume " + selectedSessionId;

                        string resolvedSessionDir = selectedSessionWorkingDirectory;
                        if (string.IsNullOrEmpty(resolvedSessionDir))
                            TryResolveCodexSessionWorkingDirectory(selectedSessionId, out resolvedSessionDir);

                        if (!string.IsNullOrEmpty(resolvedSessionDir))
                            codexLaunchDir = resolvedSessionDir.Replace("'", "''");
                    }

                    return string.Format(
                        "Set-Location -LiteralPath '{0}'; " +
                        "$codexExe = (Get-Command codex -ErrorAction SilentlyContinue).Source; " +
                        "if (-not $codexExe) {{ " +
                        "$codexExe = Get-ChildItem -Path (Join-Path $env:LOCALAPPDATA 'OpenAI\\Codex\\bin\\*\\codex.exe') " +
                        "-File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | " +
                        "Select-Object -First 1 -ExpandProperty FullName; " +
                        "}}; " +
                        "if (-not $codexExe) {{ throw 'Codex CLI was not found. Install or update the Codex app, then restart SOLIDWORKS.' }}; " +
                        "$swKey = Get-ChildItem 'HKLM:\\SOFTWARE\\SolidWorks' -ErrorAction SilentlyContinue | " +
                        "Where-Object PSChildName -Match '^SOLIDWORKS \\d{{4}}$' | Sort-Object PSChildName -Descending | Select-Object -First 1; " +
                        "if ($swKey) {{ " +
                        "$swSetup = Get-ItemProperty -LiteralPath (Join-Path $swKey.PSPath 'Setup') -ErrorAction SilentlyContinue; " +
                        "$env:SWPILOT_SOLIDWORKS_DIR = $swSetup.'SolidWorks Folder'; " +
                        "}}; " +
                        "& $codexExe -s danger-full-access -a never{1}",
                        codexLaunchDir,
                        codexResumeArg);
                case "cmd":
                    return string.Format("Set-Location -LiteralPath '{0}'; cmd", safeDir);
                default:         return string.Format("Set-Location -LiteralPath '{0}'", safeDir);
            }
        }

        private static string BuildPsReadLineWhiteCommand()
        {
            return "if (Get-Command Set-PSReadLineOption -ErrorAction SilentlyContinue) { " +
                   "$c=@{Command='White';String='White';Number='White';Operator='White';" +
                   "Variable='White';Parameter='White';Type='White';Keyword='White';Member='White'}; " +
                   "Set-PSReadLineOption -Colors $c -ErrorAction SilentlyContinue }";
        }

        private static string BuildPsLaunchCommandLine()
        {
            string script = BuildPsReadLineWhiteCommand() + "; Clear-Host";
            string encoded = Convert.ToBase64String(Encoding.Unicode.GetBytes(script));
            return "powershell.exe -NoLogo -NoExit -EncodedCommand " + encoded;
        }

        private void EmbedWindow(IntPtr hwnd)
        {
            psHwnd = hwnd;

            int style = GetWindowLong(hwnd, GWL_STYLE);
            style &= ~(WS_POPUP | WS_CAPTION | WS_THICKFRAME | WS_MINIMIZE | WS_MAXIMIZE | WS_SYSMENU);
            style |= WS_CHILD | WS_VISIBLE;
            SetWindowLong(hwnd, GWL_STYLE, style);

            int exStyle = GetWindowLong(hwnd, GWL_EXSTYLE);
            exStyle &= ~WS_EX_APPWINDOW;
            exStyle |= WS_EX_TOOLWINDOW;
            SetWindowLong(hwnd, GWL_EXSTYLE, exStyle);

            SetParent(hwnd, terminalPanel.Handle);

            SetWindowPos(hwnd, IntPtr.Zero,
                0, 0, terminalPanel.Width, terminalPanel.Height,
                SWP_FRAMECHANGED | SWP_SHOWWINDOW);

            terminalPanel.Resize += (s, e) => ResizeEmbeddedWindow();
        }

        private void ResizeEmbeddedWindow()
        {
            if (psHwnd != IntPtr.Zero && terminalPanel.Width > 0 && terminalPanel.Height > 0)
                MoveWindow(psHwnd, 0, 0, terminalPanel.Width, terminalPanel.Height, true);
        }

        protected override void OnHandleDestroyed(EventArgs e)
        {
            GracefulShutdown();
            if (psTerminal != null)
            {
                psTerminal.Dispose();
                psTerminal = null;
            }
            base.OnHandleDestroyed(e);
        }
    }
}

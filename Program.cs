// OcrTcpToKeyboard v2
// 重写说明：去掉 Interception 驱动依赖，使用剪贴板注入方式，
// 完整支持 Excel / Word / PowerPoint / WPS / 任意 Unicode 文本。
// 编译：.NET Framework 4.6+ 或 .NET 6+ (Windows)
// 依赖：无第三方库

using System;
using System.Runtime.InteropServices;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using System.Collections.Generic;

namespace OcrTcpToKeyboard2
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }

    // ─────────────────────────── Win32 辅助 ───────────────────────────
    static class Win32
    {
        [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] public static extern IntPtr GetFocus();
        [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        [DllImport("user32.dll")] public static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);
        [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll", CharSet = CharSet.Auto)] public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);
        [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, string lParam);
        [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
        [DllImport("kernel32.dll")] public static extern uint GetCurrentThreadId();

        public const uint EM_REPLACESEL = 0x00C2;
        public const uint WM_KEYDOWN = 0x0100;
        public const uint VK_RETURN = 0x0D;
        public const uint VK_END = 0x23;

        /// <summary>获取前台窗口中当前聚焦的子控件句柄</summary>
        public static IntPtr GetFocusedControl()
        {
            IntPtr hwndFg = GetForegroundWindow();
            if (hwndFg == IntPtr.Zero) return IntPtr.Zero;

            uint fgThread = GetWindowThreadProcessId(hwndFg, out _);
            uint myThread = GetCurrentThreadId();

            AttachThreadInput(myThread, fgThread, true);
            IntPtr hwndFocus = GetFocus();
            AttachThreadInput(myThread, fgThread, false);

            return hwndFocus != IntPtr.Zero ? hwndFocus : hwndFg;
        }

        public static string GetClass(IntPtr hWnd)
        {
            var sb = new StringBuilder(256);
            GetClassName(hWnd, sb, sb.Capacity);
            return sb.ToString();
        }

        /// <summary>判断是否为标准 Edit 控件（Notepad / 大多数 Win32 文本框）</summary>
        public static bool IsEditControl(string cls) =>
            cls.Equals("Edit", StringComparison.OrdinalIgnoreCase) ||
            cls.StartsWith("RichEdit", StringComparison.OrdinalIgnoreCase) ||
            cls.StartsWith("RICHEDIT", StringComparison.OrdinalIgnoreCase);

        /// <summary>判断是否为 Office / WPS 系列窗口</summary>
        public static bool IsOfficeWindow(string cls) =>
            cls.StartsWith("EXCEL", StringComparison.OrdinalIgnoreCase) ||
            cls.StartsWith("OpusApp", StringComparison.OrdinalIgnoreCase) ||   // Word
            cls.StartsWith("PP", StringComparison.OrdinalIgnoreCase) ||        // PowerPoint
            cls.StartsWith("wpsOffice", StringComparison.OrdinalIgnoreCase) || // WPS
            cls.StartsWith("WPSXMAIN", StringComparison.OrdinalIgnoreCase);
    }

    // ─────────────────────────── 注入策略 ───────────────────────────
    enum InjectStrategy { EditControl, Clipboard }

    static class TextInjector
    {
        /// <summary>
        /// 将文本注入到当前聚焦窗口。
        /// 策略1：Edit 控件 → EM_REPLACESEL（快速、无副作用）
        /// 策略2：其他（Office/WPS/任意）→ 剪贴板粘贴（Ctrl+V），追加回车
        /// </summary>
        public static (bool ok, string method, string detail) Inject(
            string text, bool appendEnter, int clipboardRestoreDelayMs = 600)
        {
            IntPtr hwnd = Win32.GetFocusedControl();
            string cls = hwnd != IntPtr.Zero ? Win32.GetClass(hwnd) : "";
            string topCls = Win32.GetClass(Win32.GetForegroundWindow());

            if (Win32.IsEditControl(cls))
            {
                // ── 策略1：直接消息注入，不影响剪贴板 ──
                string payload = appendEnter ? text + "\r\n" : text;
                Win32.SendMessage(hwnd, Win32.EM_REPLACESEL, (IntPtr)1, payload);
                return (true, "EM_REPLACESEL", $"hwnd=0x{hwnd:X} class={cls} top={topCls}");
            }
            else
            {
                // ── 策略2：剪贴板注入（支持全部 Unicode / 中文 / Office）──
                return InjectViaClipboard(text, appendEnter, clipboardRestoreDelayMs,
                    $"hwnd=0x{hwnd:X} class={cls} top={topCls}");
            }
        }

        private static (bool ok, string method, string detail) InjectViaClipboard(
            string text, bool appendEnter, int restoreMs, string detail)
        {
            string? oldClip = null;
            try
            {
                // 1. 保存当前剪贴板内容
                try { oldClip = Clipboard.ContainsText() ? Clipboard.GetText() : null; }
                catch { }

                // 2. 写入目标文本
                Clipboard.SetText(text, TextDataFormat.UnicodeText);
                Thread.Sleep(30); // 等待剪贴板稳定

                // 3. 发送 Ctrl+V
                SendKeys.SendWait("^v");

                // 4. 追加回车
                if (appendEnter)
                {
                    Thread.Sleep(20);
                    SendKeys.SendWait("{ENTER}");
                }

                // 5. 异步还原剪贴板（不阻塞主流程）
                if (oldClip != null)
                {
                    string restore = oldClip;
                    Task.Delay(restoreMs).ContinueWith(_ =>
                    {
                        try { Clipboard.SetText(restore, TextDataFormat.UnicodeText); }
                        catch { }
                    }, TaskScheduler.FromCurrentSynchronizationContext());
                }

                return (true, "Clipboard+Ctrl+V", detail);
            }
            catch (Exception ex)
            {
                return (false, "Clipboard+Ctrl+V FAIL", $"{detail} err={ex.Message}");
            }
        }
    }

    // ─────────────────────────── TCP 服务 ───────────────────────────
    class TcpListenerService : IDisposable
    {
        public event Action<string>? Log;
        public event Action<string, bool>? TextReceived; // (text, ok)

        private TcpListener? _listener;
        private CancellationTokenSource? _cts;
        private readonly int _port;
        private bool _appendEnter;

        public bool AppendEnter { get => _appendEnter; set => _appendEnter = value; }
        public bool IsRunning => _cts != null && !_cts.IsCancellationRequested;

        public TcpListenerService(int port = 62020) { _port = port; _appendEnter = true; }

        public void Start()
        {
            if (IsRunning) return;
            _cts = new CancellationTokenSource();
            _listener = new TcpListener(IPAddress.Loopback, _port);
            _listener.Start();
            Log?.Invoke($"监听 127.0.0.1:{_port}");
            Task.Run(() => AcceptLoop(_cts.Token));
        }

        public void Stop()
        {
            _cts?.Cancel();
            _listener?.Stop();
            _listener = null;
            Log?.Invoke("已停止");
        }

        private async Task AcceptLoop(CancellationToken ct)
        {
            while (!ct.IsCancellationRequested)
            {
                try
                {
                    var client = await _listener!.AcceptTcpClientAsync();
                    _ = Task.Run(() => HandleClient(client, ct), ct);
                }
                catch when (ct.IsCancellationRequested) { break; }
                catch (Exception ex) { Log?.Invoke($"Accept 错误: {ex.Message}"); }
            }
        }

        private async Task HandleClient(TcpClient client, CancellationToken ct)
        {
            var ep = client.Client.RemoteEndPoint;
            Log?.Invoke($"已连接：{ep}");
            using (client)
            using (var stream = client.GetStream())
            {
                var buf = new byte[4096];
                while (!ct.IsCancellationRequested)
                {
                    int n;
                    try { n = await stream.ReadAsync(buf, 0, buf.Length, ct); }
                    catch { break; }
                    if (n == 0) break;

                    string raw = Encoding.UTF8.GetString(buf, 0, n).Trim();
                    if (string.IsNullOrEmpty(raw))
                    {
                        Log?.Invoke("RECV: <EMPTY> → 忽略");
                        continue;
                    }

                    Log?.Invoke($"RECV: {raw}");

                    // STA 线程执行注入（SendKeys / Clipboard 必须在 STA）
                    bool ok = false;
                    Application.OpenForms[0]?.Invoke((Action)(() =>
                    {
                        var (success, method, detail) = TextInjector.Inject(raw, _appendEnter);
                        ok = success;
                        Log?.Invoke($"SEND({method}) {(success ? "OK" : "FAIL")}: {detail}");
                        TextReceived?.Invoke(raw, success);
                    }));

                    // 回复状态
                    var reply = Encoding.UTF8.GetBytes(ok ? "OK\n" : "FAIL\n");
                    try { await stream.WriteAsync(reply, 0, reply.Length, ct); }
                    catch { break; }
                }
            }
            Log?.Invoke($"连接断开：{ep}");
        }

        public void Dispose() => Stop();
    }

    // ─────────────────────────── 主窗体 ───────────────────────────
    class MainForm : Form
    {
        private readonly TcpListenerService _service = new TcpListenerService(62020);
        private Button _btnToggle = null!;
        private Button _btnTest = null!;
        private CheckBox _chkEnter = null!;
        private TextBox _txtLog = null!;
        private Label _lblStatus = null!;
        private NotifyIcon _tray = null!;
        private readonly List<string> _logLines = new();
        private const int MaxLog = 500;

        public MainForm()
        {
            BuildUI();
            _service.Log += line => AppendLog(line);
            _service.TextReceived += (text, ok) => { /* 可扩展：统计等 */ };
        }

        // ── UI 构建 ──
        void BuildUI()
        {
            Text = "OcrTcpToKeyboard v2";
            Size = new Size(640, 500);
            MinimumSize = new Size(480, 380);
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Segoe UI", 9f);

            // ── 顶部控制栏 ──
            var topPanel = new Panel { Dock = DockStyle.Top, Height = 56, Padding = new Padding(10, 8, 10, 4) };

            _btnToggle = new Button
            {
                Text = "▶  启动监听",
                Width = 120, Height = 36,
                Location = new Point(10, 10),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(0, 120, 215),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            _btnToggle.FlatAppearance.BorderSize = 0;
            _btnToggle.Click += BtnToggle_Click;

            _btnTest = new Button
            {
                Text = "发送测试文字",
                Width = 110, Height = 36,
                Location = new Point(140, 10),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(240, 240, 240),
                ForeColor = Color.FromArgb(30, 30, 30),
                Cursor = Cursors.Hand
            };
            _btnTest.FlatAppearance.BorderSize = 1;
            _btnTest.Click += BtnTest_Click;

            _chkEnter = new CheckBox
            {
                Text = "末尾追加回车 (Enter)",
                Checked = true,
                AutoSize = true,
                Location = new Point(264, 18),
                Cursor = Cursors.Hand
            };
            _chkEnter.CheckedChanged += (_, __) => _service.AppendEnter = _chkEnter.Checked;

            _lblStatus = new Label
            {
                Text = "● 未运行",
                AutoSize = true,
                Location = new Point(430, 20),
                ForeColor = Color.Gray,
                Font = new Font("Segoe UI", 9f)
            };

            topPanel.Controls.AddRange(new Control[] { _btnToggle, _btnTest, _chkEnter, _lblStatus });

            // ── 说明标签 ──
            var lblInfo = new Label
            {
                Text = "支持注入：Excel / Word / PowerPoint / WPS / Notepad / 任意文本框（含中文/Unicode）",
                Dock = DockStyle.Top,
                Height = 24,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(12, 0, 0, 0),
                ForeColor = Color.FromArgb(80, 80, 80),
                Font = new Font("Segoe UI", 8.5f),
                BackColor = Color.FromArgb(248, 248, 248)
            };

            // ── 日志区 ──
            _txtLog = new TextBox
            {
                Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(30, 30, 30),
                ForeColor = Color.FromArgb(200, 240, 200),
                Font = new Font("Consolas", 8.5f),
                BorderStyle = BorderStyle.None
            };

            var btnClearLog = new Button
            {
                Text = "清空日志",
                Dock = DockStyle.Bottom,
                Height = 26,
                FlatStyle = FlatStyle.Flat,
                ForeColor = Color.Gray,
                BackColor = Color.FromArgb(240, 240, 240)
            };
            btnClearLog.FlatAppearance.BorderSize = 0;
            btnClearLog.Click += (_, __) => { _logLines.Clear(); _txtLog.Clear(); };

            Controls.Add(_txtLog);
            Controls.Add(btnClearLog);
            Controls.Add(lblInfo);
            Controls.Add(topPanel);

            // ── 系统托盘 ──
            _tray = new NotifyIcon
            {
                Icon = SystemIcons.Application,
                Text = "OcrTcpToKeyboard v2",
                Visible = true
            };
            var trayMenu = new ContextMenuStrip();
            trayMenu.Items.Add("显示窗口", null, (_, __) => { Show(); WindowState = FormWindowState.Normal; });
            trayMenu.Items.Add("退出", null, (_, __) => { _service.Stop(); _tray.Visible = false; Application.Exit(); });
            _tray.ContextMenuStrip = trayMenu;
            _tray.DoubleClick += (_, __) => { Show(); WindowState = FormWindowState.Normal; };

            FormClosing += (_, e) => { e.Cancel = true; Hide(); };
            Resize += (_, __) => { if (WindowState == FormWindowState.Minimized) Hide(); };
        }

        void BtnToggle_Click(object? sender, EventArgs e)
        {
            if (_service.IsRunning)
            {
                _service.Stop();
                _btnToggle.Text = "▶  启动监听";
                _btnToggle.BackColor = Color.FromArgb(0, 120, 215);
                _lblStatus.Text = "● 未运行";
                _lblStatus.ForeColor = Color.Gray;
            }
            else
            {
                try
                {
                    _service.Start();
                    _btnToggle.Text = "■  停止监听";
                    _btnToggle.BackColor = Color.FromArgb(196, 43, 28);
                    _lblStatus.Text = "● 运行中  127.0.0.1:62020";
                    _lblStatus.ForeColor = Color.FromArgb(0, 150, 60);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"启动失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        void BtnTest_Click(object? sender, EventArgs e)
        {
            var dlg = new InputDialog("输入测试文字", "Hello World 你好世界 テスト");
            if (dlg.ShowDialog(this) == DialogResult.OK && !string.IsNullOrEmpty(dlg.Value))
            {
                AppendLog($"[手动测试] 注入: {dlg.Value}");
                var (ok, method, detail) = TextInjector.Inject(dlg.Value, _chkEnter.Checked);
                AppendLog($"SEND({method}) {(ok ? "OK" : "FAIL")}: {detail}");
            }
        }

        void AppendLog(string line)
        {
            if (InvokeRequired) { Invoke((Action<string>)AppendLog, line); return; }
            string ts = DateTime.Now.ToString("HH:mm:ss.fff");
            string entry = $"[{ts}] {line}";
            _logLines.Add(entry);
            if (_logLines.Count > MaxLog) _logLines.RemoveAt(0);
            _txtLog.Lines = _logLines.ToArray();
            _txtLog.SelectionStart = _txtLog.Text.Length;
            _txtLog.ScrollToCaret();
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            _service.Dispose();
            _tray.Visible = false;
            base.OnFormClosed(e);
        }
    }

    // ─────────────────────────── 简单输入对话框 ───────────────────────────
    class InputDialog : Form
    {
        private TextBox _tb = null!;
        private Button _ok = null!, _cancel = null!;
        public string Value => _tb.Text;

        public InputDialog(string title, string defaultText = "")
        {
            Text = title; Size = new Size(420, 130);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = MinimizeBox = false; StartPosition = FormStartPosition.CenterParent;

            _tb = new TextBox { Text = defaultText, Location = new Point(12, 12), Width = 380, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
            _ok = new Button { Text = "确定", DialogResult = DialogResult.OK, Location = new Point(220, 44), Width = 80 };
            _cancel = new Button { Text = "取消", DialogResult = DialogResult.Cancel, Location = new Point(312, 44), Width = 80 };
            AcceptButton = _ok; CancelButton = _cancel;
            Controls.AddRange(new Control[] { _tb, _ok, _cancel });
        }
    }
}

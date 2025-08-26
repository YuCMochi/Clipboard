using System;
using System.Windows.Forms;
using System.Drawing;
using System.IO;
using System.Diagnostics;
using System.Threading;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace ClipboardApp
{
    public partial class MainForm : Form
    {
        private NotifyIcon trayIcon;
        private ContextMenuStrip trayMenu;
        private ToolStripMenuItem toggleMonitorMenuItem;
        private ToolStripMenuItem exitMenuItem;
        private bool isMonitoring = true;
        private string lastOpenedPath = "";
        private Icon iconOn;
        private Icon iconOff;
        private static Mutex? singleInstanceMutex;

        // Startup registry config
        private const string RunRegPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
        private const string RunValueName = "ClipboardApp";

        // Clipboard event listener
        private const int WM_CLIPBOARDUPDATE = 0x031D;
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool AddClipboardFormatListener(IntPtr hwnd);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool RemoveClipboardFormatListener(IntPtr hwnd);

        public MainForm()
        {
            LoadIcons();
            InitializeComponent();
            InitTray();
            this.WindowState = FormWindowState.Minimized;
            this.ShowInTaskbar = false;
            this.Visible = false;
            var _ = this.Handle; // 確保建立視窗 Handle 以註冊剪貼簿事件
        }

        [STAThread]
        static void Main()
        {
            bool createdNew;
            singleInstanceMutex = new Mutex(true, "ClipboardApp_SingleInstance_Mutex", out createdNew);
            if (!createdNew)
            {
                // 已有實例在執行，直接結束
                return;
            }
            ApplicationConfiguration.Initialize();
            Application.Run(new MainForm());
        }

        private Icon LoadBestIcon(string[] paths)
        {
            foreach (var path in paths)
            {
                string absPath = Path.GetFullPath(path, AppDomain.CurrentDomain.BaseDirectory);
                if (File.Exists(absPath))
                {
                    return new Icon(absPath);
                }
            }
            return SystemIcons.Application;
        }

        private void LoadIcons()
        {
            string exeDir = AppDomain.CurrentDomain.BaseDirectory;
            string iconDir = Path.Combine(exeDir, "icon");
            iconOn = LoadBestIcon(new string[] {
                Path.Combine(iconDir, "on", "32x32.ico"),
                Path.Combine(iconDir, "on", "16x16.ico")
            });
            iconOff = LoadBestIcon(new string[] {
                Path.Combine(iconDir, "off", "32x32.ico"),
                Path.Combine(iconDir, "off", "16x16.ico")
            });
        }

        private void InitTray()
        {
            trayMenu = new ContextMenuStrip();
            var startupItem = new ToolStripMenuItem("開機自動啟動") { Checked = IsStartupEnabled() };
            startupItem.Click += (s, e) =>
            {
                bool next = !startupItem.Checked;
                SetStartup(next);
                startupItem.Checked = next;
            };
            toggleMonitorMenuItem = new ToolStripMenuItem("暫停監控", null, OnToggleMonitor);
            exitMenuItem = new ToolStripMenuItem("退出", null, OnExit);
            trayMenu.Items.Add(startupItem);
            trayMenu.Items.Add(toggleMonitorMenuItem);
            trayMenu.Items.Add(exitMenuItem);

            trayIcon = new NotifyIcon();
            trayIcon.Text = "Clipboard 路徑自動開啟";
            trayIcon.Icon = isMonitoring ? iconOn : iconOff;
            trayIcon.ContextMenuStrip = trayMenu;
            trayIcon.Visible = true;
            trayIcon.MouseClick += TrayIcon_MouseClick;
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            try { AddClipboardFormatListener(this.Handle); } catch { }
        }

        protected override void OnHandleDestroyed(EventArgs e)
        {
            try { RemoveClipboardFormatListener(this.Handle); } catch { }
            base.OnHandleDestroyed(e);
        }

        private bool IsValidPath(string path)
        {
            // 支援多語言路徑，判斷是否為存在的檔案或資料夾
            return Directory.Exists(path) || File.Exists(path);
        }

        private string NormalizePath(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return string.Empty;
            string text = raw.Trim().Trim('"');

            if (text.StartsWith("file://", StringComparison.OrdinalIgnoreCase))
            {
                try { text = new Uri(text).LocalPath; } catch { }
            }

            text = text.Replace('/', '\\');
            text = Environment.ExpandEnvironmentVariables(text);
            return text.Trim();
        }

        private void OpenPath(string path)
        {
            try
            {
                Process.Start(new ProcessStartInfo()
                {
                    FileName = path,
                    UseShellExecute = true
                });
                lastOpenedPath = path;
            }
            catch { /* 忽略開啟失敗 */ }
        }

        private void TrayIcon_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                OnToggleMonitor(sender, e);
            }
        }

        private void OnToggleMonitor(object sender, EventArgs e)
        {
            isMonitoring = !isMonitoring;
            toggleMonitorMenuItem.Text = isMonitoring ? "暫停監控" : "恢復監控";
            trayIcon.Icon = isMonitoring ? iconOn : iconOff;
            // 不再顯示 BalloonTip
        }

        private void OnExit(object sender, EventArgs e)
        {
            trayIcon.Visible = false;
            Application.Exit();
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            trayIcon.Visible = false;
            base.OnFormClosing(e);
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_CLIPBOARDUPDATE && isMonitoring)
            {
                try
                {
                    if (Clipboard.ContainsText())
                    {
                        string text = NormalizePath(Clipboard.GetText());
                        if (!string.IsNullOrEmpty(text) && IsValidPath(text))
                        {
                            OpenPath(text);
                        }
                    }
                }
                catch { }
            }
            base.WndProc(ref m);
        }

        private bool IsStartupEnabled()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(RunRegPath, writable: false);
                var val = key?.GetValue(RunValueName) as string;
                string exe = Application.ExecutablePath;
                return !string.IsNullOrEmpty(val) && val.Trim('"').Equals(exe, StringComparison.OrdinalIgnoreCase);
            }
            catch { return false; }
        }

        private void SetStartup(bool enable)
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(RunRegPath, writable: true);
                if (key == null) return;

                string exe = Application.ExecutablePath;
                if (enable)
                {
                    key.SetValue(RunValueName, $"\"{exe}\"");
                }
                else
                {
                    key.DeleteValue(RunValueName, throwOnMissingValue: false);
                }
            }
            catch { }
        }

        private void InitializeComponent() { /* 無設計師元件，空實作 */ }
    }
} 
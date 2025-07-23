using System;
using System.Windows.Forms;
using System.Drawing;
using System.IO;
using System.Diagnostics;

public partial class MainForm : Form
{
    private NotifyIcon trayIcon;
    private ContextMenuStrip trayMenu;
    private ToolStripMenuItem toggleMonitorMenuItem;
    private ToolStripMenuItem exitMenuItem;
    private bool isMonitoring = true;
    private Timer clipboardTimer;
    private string lastClipboardText = "";

    public MainForm()
    {
        InitializeComponent();
        InitTray();
        InitClipboardMonitor();
        this.WindowState = FormWindowState.Minimized;
        this.ShowInTaskbar = false;
        this.Visible = false;
    }

    private void InitTray()
    {
        trayMenu = new ContextMenuStrip();
        toggleMonitorMenuItem = new ToolStripMenuItem("暫停監控", null, OnToggleMonitor);
        exitMenuItem = new ToolStripMenuItem("退出", null, OnExit);
        trayMenu.Items.Add(toggleMonitorMenuItem);
        trayMenu.Items.Add(exitMenuItem);

        trayIcon = new NotifyIcon();
        trayIcon.Text = "Clipboard 路徑自動開啟";
        trayIcon.Icon = SystemIcons.Application;
        trayIcon.ContextMenuStrip = trayMenu;
        trayIcon.Visible = true;
        trayIcon.MouseClick += TrayIcon_MouseClick;
    }

    private void InitClipboardMonitor()
    {
        clipboardTimer = new Timer();
        clipboardTimer.Interval = 800; // 每0.8秒檢查一次
        clipboardTimer.Tick += ClipboardTimer_Tick;
        clipboardTimer.Start();
    }

    private void ClipboardTimer_Tick(object sender, EventArgs e)
    {
        if (!isMonitoring) return;
        try
        {
            if (Clipboard.ContainsText())
            {
                string text = Clipboard.GetText().Trim();
                if (text != lastClipboardText && IsValidPath(text))
                {
                    lastClipboardText = text;
                    OpenPath(text);
                }
            }
        }
        catch { /* 忽略剪貼簿存取例外 */ }
    }

    private bool IsValidPath(string path)
    {
        // 支援多語言路徑，判斷是否為存在的檔案或資料夾
        return Directory.Exists(path) || File.Exists(path);
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
        trayIcon.ShowBalloonTip(1000, "Clipboard 路徑自動開啟", isMonitoring ? "已恢復監控" : "已暫停監控", ToolTipIcon.Info);
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
} 
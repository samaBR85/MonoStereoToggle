using System.Diagnostics;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace MonoStereoToggle;

public sealed class MainForm : Form
{
    // ─── Win32 ────────────────────────────────────────────────────────────────

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    private static extern IntPtr SendMessageTimeout(IntPtr hWnd, uint Msg, UIntPtr wParam,
        string? lParam, uint fuFlags, uint uTimeout, out UIntPtr lpdwResult);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool SystemParametersInfo(uint uAction, uint uParam,
        ref ACCESSTIMEOUT pvParam, uint fWinIni);

    [DllImport("dwmapi.dll")]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr,
        ref int attrValue, int attrSize);

    [DllImport("user32.dll")]
    private static extern bool DestroyIcon(IntPtr hIcon);

    // ── Global hotkey ─────────────────────────────────────────────────────────
    [DllImport("user32.dll")]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

    [DllImport("user32.dll")]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    [StructLayout(LayoutKind.Sequential)]
    private struct ACCESSTIMEOUT { public uint cbSize, dwFlags, iTimeOutMSec; }

    private static readonly IntPtr HWND_BROADCAST   = new(0xFFFF);
    private const uint WM_SETTINGCHANGE             = 0x001A;
    private const uint SMTO_ABORTIFHUNG             = 0x0002;
    private const uint SMTO_NOTIMEOUTIFHUNG         = 0x0008;
    private const uint SPI_GETACCESSTIMEOUT         = 0x003E;
    private const uint SPI_SETACCESSTIMEOUT         = 0x003D;
    private const uint SPIF_UPDATEINIFILE           = 0x0001;
    private const uint SPIF_SENDCHANGE              = 0x0002;

    private const uint MOD_ALT      = 0x0001;
    private const uint MOD_CONTROL  = 0x0002;
    private const uint MOD_SHIFT    = 0x0004;
    private const uint MOD_WIN      = 0x0008;
    private const uint MOD_NOREPEAT = 0x4000;
    private const uint WM_HOTKEY    = 0x0312;
    private const int  HOTKEY_ID    = 0x4D53; // arbitrary non-zero ID

    // ─── Core Audio COM ───────────────────────────────────────────────────────

    private enum EDataFlow { eRender, eCapture, eAll }
    private enum ERole     { eConsole, eMultimedia, eCommunications }

    [ComImport, Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
    private class CMMDeviceEnumerator { }

    [ComImport, Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IMMDeviceEnumerator
    {
        [PreserveSig] int EnumAudioEndpoints(EDataFlow f, uint mask,
            [MarshalAs(UnmanagedType.Interface)] out IMMDeviceCollection c);
        [PreserveSig] int GetDefaultAudioEndpoint(EDataFlow f, ERole r,
            [MarshalAs(UnmanagedType.Interface)] out IMMDevice d);
        [PreserveSig] int GetDevice([MarshalAs(UnmanagedType.LPWStr)] string id,
            [MarshalAs(UnmanagedType.Interface)] out IMMDevice d);
        void _RegisterEndpointNotificationCallback();
        void _UnregisterEndpointNotificationCallback();
    }

    [ComImport, Guid("0BD7A1BE-7A1A-44DB-8397-CC5392387B5E"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IMMDeviceCollection
    {
        [PreserveSig] int GetCount(out uint n);
        [PreserveSig] int Item(uint i, [MarshalAs(UnmanagedType.Interface)] out IMMDevice d);
    }

    [ComImport, Guid("D666063F-1587-4E43-81F1-B948E807363F"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IMMDevice
    {
        [PreserveSig] int Activate(ref Guid iid, uint ctx, IntPtr p,
            [MarshalAs(UnmanagedType.IUnknown)] out object pp);
        [PreserveSig] int OpenPropertyStore(uint access,
            [MarshalAs(UnmanagedType.Interface)] out IPropertyStore pp);
        [PreserveSig] int GetId([MarshalAs(UnmanagedType.LPWStr)] out string id);
        [PreserveSig] int GetState(out uint state);
    }

    [ComImport, Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IPropertyStore
    {
        [PreserveSig] int GetCount(out uint n);
        [PreserveSig] int GetAt(uint i, out PROPERTYKEY k);
        [PreserveSig] int GetValue(ref PROPERTYKEY k, out PROPVARIANT v);
        [PreserveSig] int SetValue(ref PROPERTYKEY k, ref PROPVARIANT v);
        [PreserveSig] int Commit();
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct PROPERTYKEY { public Guid fmtid; public uint pid; }

    [StructLayout(LayoutKind.Explicit, Size = 16)]
    private struct PROPVARIANT
    {
        [FieldOffset(0)] public ushort vt;
        [FieldOffset(8)] public IntPtr ptr;
        public string? AsString() => vt == 0x1F ? Marshal.PtrToStringUni(ptr) : null;
    }

    private static readonly PROPERTYKEY PKEY_FriendlyName = new()
    { fmtid = new Guid("a45c254e-df1c-4efd-8020-67d146a850e0"), pid = 14 };

    private const uint STGM_READ = 0;
    private const uint DEVICE_STATE_ACTIVE = 1;

    // ─── Registry ─────────────────────────────────────────────────────────────

    private const string AUDIO_KEY  = @"SOFTWARE\Microsoft\Multimedia\Audio";
    private const string MONO_VALUE = "AccessibilityMonoMixState";
    private const string ACCESS_KEY = @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Accessibility";
    private const string CONFIG_VAL = "Configuration";
    private const string MONO_FEAT  = "monoaudio";
    private const string RUN_KEY    = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";
    private const string APP_KEY    = @"SOFTWARE\MonoStereoToggle";
    private const string APP_NAME   = "MonoStereoToggle";

    // ─── State ────────────────────────────────────────────────────────────────

    private bool _isMono;           // mirrors registry AccessibilityMonoMixState
    private bool _busy;             // prevent double-click during service restart
    private readonly bool _startHidden;
    private bool _initialShowDone;
    private bool _exitRequested;

    // Hotkey
    private bool _recordingHotkey;
    private uint _hotkeyModifiers;
    private uint _hotkeyVk;
    private bool _hotkeyRegistered;

    // Overlay
    private OverlayForm? _overlay;

    // ─── Controls ─────────────────────────────────────────────────────────────

    private readonly NotifyIcon        _tray;
    private readonly ContextMenuStrip  _trayMenu;
    private readonly ToolStripMenuItem _trayToggleItem;
    private readonly Button            _btnToggle;
    private readonly Label             _lblStatus;
    private readonly ListBox           _listDevices;
    private readonly CheckBox          _chkStartup;
    private readonly CheckBox          _chkTray;
    private readonly TextBox           _txtHotkey;
    private readonly Button            _btnSetHotkey;
    private readonly Button            _btnClearHotkey;
    private readonly Icon              _iconMono;
    private readonly Icon              _iconStereo;

    private List<AudioDevice> _devices = new();

    // ─── Constructor ──────────────────────────────────────────────────────────

    public MainForm(bool startHidden)
    {
        // ── Read current system state from registry ───────────────────────────
        _isMono      = RegGet<int>(AUDIO_KEY, MONO_VALUE) == 1;
        _startHidden = startHidden || RegGet<int>(APP_KEY, "StartInTray") == 1;

        _iconMono   = MakeIcon(Color.FromArgb(0, 103, 192), "M");
        _iconStereo = MakeIcon(Color.FromArgb(75, 75, 75),  "S");

        // ── Form ──────────────────────────────────────────────────────────────
        Text            = APP_NAME;
        Size            = new Size(390, 530);
        MinimumSize     = new Size(340, 480);   // resizable but not too small
        FormBorderStyle = FormBorderStyle.Sizable;
        MaximizeBox     = true;
        StartPosition   = FormStartPosition.CenterScreen;
        BackColor       = Color.FromArgb(28, 28, 28);
        ForeColor       = Color.White;
        Font            = new Font("Segoe UI", 10f);

        // ── Title ─────────────────────────────────────────────────────────────
        var lblTitle = MkLabel(Strings.AppTitle,
            new Font(new FontFamily("Segoe UI Light"), 16f), Color.White, new Point(20, 16));
        lblTitle.Anchor = AnchorStyles.Top | AnchorStyles.Left;
        Controls.Add(lblTitle);

        // ── Status ────────────────────────────────────────────────────────────
        _lblStatus = MkLabel("", new Font("Segoe UI", 9f),
            Color.FromArgb(140, 140, 140), new Point(22, 54));
        _lblStatus.Anchor = AnchorStyles.Top | AnchorStyles.Left;
        Controls.Add(_lblStatus);

        // ── Toggle button ─────────────────────────────────────────────────────
        _btnToggle = new Button
        {
            Location = new Point(20, 76),
            Size     = new Size(ClientSize.Width - 40, 88),
            Anchor   = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            FlatStyle = FlatStyle.Flat,
            Font      = new Font("Segoe UI", 24f, FontStyle.Bold),
            ForeColor = Color.White,
            Cursor    = Cursors.Hand,
            UseVisualStyleBackColor = false
        };
        _btnToggle.FlatAppearance.BorderSize         = 0;
        _btnToggle.FlatAppearance.MouseOverBackColor = Color.Transparent;
        _btnToggle.FlatAppearance.MouseDownBackColor = Color.Transparent;
        _btnToggle.Paint     += PaintButton;
        _btnToggle.Click     += (_, _) => ForceToggle();
        _btnToggle.SizeChanged += (_, _) => RoundRegion(_btnToggle, 10);
        RoundRegion(_btnToggle, 10);
        Controls.Add(_btnToggle);

        // ── Separator 1 ───────────────────────────────────────────────────────
        var sep1 = new Panel
        {
            Location  = new Point(20, 178),
            Size      = new Size(ClientSize.Width - 40, 1),
            Anchor    = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            BackColor = Color.FromArgb(52, 52, 52)
        };
        Controls.Add(sep1);

        // ── Devices label ─────────────────────────────────────────────────────
        var lblDev = MkLabel(Strings.DevicesLabel,
            new Font("Segoe UI", 8.5f), Color.FromArgb(160, 160, 160), new Point(20, 188));
        lblDev.Anchor = AnchorStyles.Top | AnchorStyles.Left;
        Controls.Add(lblDev);

        // ── Device list ───────────────────────────────────────────────────────
        _listDevices = new ListBox
        {
            Location      = new Point(20, 208),
            Size          = new Size(ClientSize.Width - 40, ClientSize.Height - 408),
            Anchor        = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
            BackColor     = Color.FromArgb(38, 38, 38),
            ForeColor     = Color.FromArgb(210, 210, 210),
            BorderStyle   = BorderStyle.FixedSingle,
            Font          = new Font("Segoe UI", 9f),
            SelectionMode = SelectionMode.None,
            DrawMode      = DrawMode.OwnerDrawFixed,
            ItemHeight    = 24
        };
        _listDevices.DrawItem += DrawDeviceItem;
        Controls.Add(_listDevices);

        // ── Device hint ───────────────────────────────────────────────────────
        var lblDeviceHint = MkLabel(Strings.DeviceHint,
            new Font("Segoe UI", 7.5f), Color.FromArgb(80, 80, 80),
            new Point(20, ClientSize.Height - 218));
        lblDeviceHint.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
        Controls.Add(lblDeviceHint);

        // ── Separator 2 ───────────────────────────────────────────────────────
        var sep2 = new Panel
        {
            Location  = new Point(20, ClientSize.Height - 200),
            Size      = new Size(ClientSize.Width - 40, 1),
            Anchor    = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            BackColor = Color.FromArgb(52, 52, 52)
        };
        Controls.Add(sep2);

        // ── Startup checkbox ──────────────────────────────────────────────────
        _chkStartup = MkCheckBox(Strings.ChkStartup,
            new Point(20, ClientSize.Height - 186),
            IsStartupTaskRegistered());
        _chkStartup.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
        _chkStartup.CheckedChanged += (_, _) => ApplyStartup(_chkStartup.Checked);
        Controls.Add(_chkStartup);

        // ── Tray checkbox ─────────────────────────────────────────────────────
        _chkTray = MkCheckBox(Strings.ChkTray,
            new Point(20, ClientSize.Height - 154),
            RegGet<int>(APP_KEY, "StartInTray") == 1);
        _chkTray.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
        _chkTray.CheckedChanged += (_, _) =>
            RegSet(APP_KEY, "StartInTray", _chkTray.Checked ? 1 : 0);
        Controls.Add(_chkTray);

        // ── Separator 3 ───────────────────────────────────────────────────────
        var sep3 = new Panel
        {
            Location  = new Point(20, ClientSize.Height - 122),
            Size      = new Size(ClientSize.Width - 40, 1),
            Anchor    = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            BackColor = Color.FromArgb(52, 52, 52)
        };
        Controls.Add(sep3);

        // ── Hotkey row ────────────────────────────────────────────────────────
        var lblHk = MkLabel(Strings.HotkeyLabel, new Font("Segoe UI", 9f),
            Color.FromArgb(160, 160, 160), new Point(20, ClientSize.Height - 103));
        lblHk.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
        Controls.Add(lblHk);

        _txtHotkey = new TextBox
        {
            Location  = new Point(74, ClientSize.Height - 106),
            Size      = new Size(ClientSize.Width - 234, 23),
            Anchor    = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            ReadOnly  = true,
            BackColor = Color.FromArgb(38, 38, 38),
            ForeColor = Color.FromArgb(200, 200, 200),
            Font      = new Font("Segoe UI", 9f),
            BorderStyle = BorderStyle.FixedSingle,
            Text      = Strings.HotkeyNone,
            Cursor    = Cursors.Default,
            TabStop   = false
        };
        Controls.Add(_txtHotkey);

        _btnSetHotkey = new Button
        {
            Location  = new Point(ClientSize.Width - 156, ClientSize.Height - 108),
            Size      = new Size(68, 26),
            Anchor    = AnchorStyles.Bottom | AnchorStyles.Right,
            Text      = Strings.HotkeySet,
            FlatStyle = FlatStyle.Flat,
            Font      = new Font("Segoe UI", 8.5f),
            ForeColor = Color.FromArgb(200, 200, 200),
            BackColor = Color.FromArgb(52, 52, 52),
            Cursor    = Cursors.Hand,
        };
        _btnSetHotkey.FlatAppearance.BorderColor = Color.FromArgb(80, 80, 80);
        _btnSetHotkey.Click += OnSetHotkeyClick;
        Controls.Add(_btnSetHotkey);

        _btnClearHotkey = new Button
        {
            Location  = new Point(ClientSize.Width - 84, ClientSize.Height - 108),
            Size      = new Size(64, 26),
            Anchor    = AnchorStyles.Bottom | AnchorStyles.Right,
            Text      = Strings.HotkeyClear,
            FlatStyle = FlatStyle.Flat,
            Font      = new Font("Segoe UI", 8.5f),
            ForeColor = Color.FromArgb(160, 80, 80),
            BackColor = Color.FromArgb(52, 52, 52),
            Cursor    = Cursors.Hand,
        };
        _btnClearHotkey.FlatAppearance.BorderColor = Color.FromArgb(80, 80, 80);
        _btnClearHotkey.Click += (_, _) => ClearHotkey();
        Controls.Add(_btnClearHotkey);

        // ── Footer ────────────────────────────────────────────────────────────
        var footer = MkLabel(Strings.Footer,
            new Font("Segoe UI", 7.5f), Color.FromArgb(70, 70, 70),
            new Point(20, ClientSize.Height - 52));
        footer.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
        Controls.Add(footer);

        // ── Key capture setup ─────────────────────────────────────────────────
        KeyPreview = true;
        KeyDown   += OnFormKeyDown;

        // ── Tray ──────────────────────────────────────────────────────────────
        _trayToggleItem = new ToolStripMenuItem
            { Font = new Font("Segoe UI", 9f, FontStyle.Bold) };
        _trayToggleItem.Click += (_, _) => ForceToggle();

        _trayMenu = new ContextMenuStrip { Renderer = new DarkMenuRenderer() };
        _trayMenu.Items.Add(_trayToggleItem);
        _trayMenu.Items.Add(new ToolStripSeparator());
        _trayMenu.Items.Add(Strings.TrayOpen,  null, (_, _) => RestoreWindow());
        _trayMenu.Items.Add(Strings.TrayClose, null, (_, _) => ExitApp());

        _tray = new NotifyIcon { ContextMenuStrip = _trayMenu, Visible = true };
        _tray.DoubleClick += (_, _) => RestoreWindow();

        LoadDevices();
        ApplyState();
    }

    // ─── Overrides ────────────────────────────────────────────────────────────

    protected override void SetVisibleCore(bool value)
    {
        if (!_initialShowDone)
        {
            _initialShowDone = true;
            if (_startHidden) { CreateHandle(); base.SetVisibleCore(false); return; }
        }
        base.SetVisibleCore(value);
    }

    protected override void OnHandleCreated(EventArgs e)
    {
        base.OnHandleCreated(e);
        int v = 1;
        if (DwmSetWindowAttribute(Handle, 20, ref v, sizeof(int)) != 0)
            DwmSetWindowAttribute(Handle, 19, ref v, sizeof(int));
        RegisterSavedHotkey();
    }

    protected override void WndProc(ref Message m)
    {
        if (m.Msg == WM_HOTKEY && m.WParam.ToInt32() == HOTKEY_ID)
            ForceToggle();
        base.WndProc(ref m);
    }

    protected override void OnDeactivate(EventArgs e)
    {
        base.OnDeactivate(e);
        if (_recordingHotkey) CancelHotkeyRecording();
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        if (!_exitRequested && e.CloseReason == CloseReason.UserClosing && _chkTray.Checked)
        {
            e.Cancel = true; Hide(); return;
        }
        _tray.Visible = false;
        base.OnFormClosing(e);
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            if (_hotkeyRegistered) UnregisterHotKey(Handle, HOTKEY_ID);
            _tray.Dispose(); _trayMenu.Dispose();
            _iconMono.Dispose(); _iconStereo.Dispose();
        }
        base.Dispose(disposing);
    }

    // ─── Device listing ───────────────────────────────────────────────────────

    private record AudioDevice(string Id, string Name, bool IsDefault);

    private void LoadDevices()
    {
        _devices = EnumerateRenderDevices();
        _listDevices.Items.Clear();
        foreach (var d in _devices) _listDevices.Items.Add(d);
    }

    private static List<AudioDevice> EnumerateRenderDevices()
    {
        var result = new List<AudioDevice>();
        try
        {
            var enumerator = (IMMDeviceEnumerator)new CMMDeviceEnumerator();
            string defaultId = "";
            try
            {
                enumerator.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia, out var def);
                def.GetId(out defaultId);
                Marshal.ReleaseComObject(def);
            }
            catch { }

            enumerator.EnumAudioEndpoints(EDataFlow.eRender, DEVICE_STATE_ACTIVE, out var col);
            col.GetCount(out uint count);
            for (uint i = 0; i < count; i++)
            {
                col.Item(i, out var dev);
                dev.GetId(out string id);
                string name = GetFriendlyName(dev) ?? id;
                result.Add(new AudioDevice(id, name, id == defaultId));
                Marshal.ReleaseComObject(dev);
            }
            Marshal.ReleaseComObject(col);
            Marshal.ReleaseComObject(enumerator);
        }
        catch { }
        return result;
    }

    private static string? GetFriendlyName(IMMDevice dev)
    {
        try
        {
            dev.OpenPropertyStore(STGM_READ, out var store);
            var key = PKEY_FriendlyName;
            store.GetValue(ref key, out PROPVARIANT v);
            var name = v.AsString();
            Marshal.ReleaseComObject(store);
            return name;
        }
        catch { return null; }
    }

    private void DrawDeviceItem(object? sender, DrawItemEventArgs e)
    {
        if (e.Index < 0 || e.Index >= _devices.Count) return;
        var dev  = _devices[e.Index];
        var g    = e.Graphics;
        var rect = e.Bounds;

        using (var bg = new SolidBrush(Color.FromArgb(38, 38, 38)))
            g.FillRectangle(bg, rect);

        var dotColor = dev.IsDefault ? Color.FromArgb(0, 180, 100) : Color.FromArgb(65, 65, 65);
        using (var dot = new SolidBrush(dotColor))
            g.FillEllipse(dot, rect.X + 8, rect.Y + 8, 8, 8);

        var nameColor = dev.IsDefault ? Color.White : Color.FromArgb(155, 155, 155);
        using var nameFont = dev.IsDefault
            ? new Font("Segoe UI", 9f, FontStyle.Bold)
            : new Font("Segoe UI", 9f);
        using (var nameBrush = new SolidBrush(nameColor))
            g.DrawString(dev.Name, nameFont, nameBrush, new PointF(rect.X + 24, rect.Y + 4));

        if (dev.IsDefault)
        {
            var tw = g.MeasureString(dev.Name, nameFont).Width;
            using var tagFont  = new Font("Segoe UI", 7.5f);
            using var tagBrush = new SolidBrush(Color.FromArgb(0, 150, 80));
            g.DrawString(Strings.DeviceDefault, tagFont, tagBrush,
                new PointF(rect.X + 24 + tw + 4, rect.Y + 6));
        }

        using var divider = new Pen(Color.FromArgb(50, 50, 50));
        g.DrawLine(divider, rect.Left, rect.Bottom - 1, rect.Right, rect.Bottom - 1);
    }

    // ─── Toggle ───────────────────────────────────────────────────────────────

    /// <summary>
    /// Toggle audio mode. The app already runs as admin (manifest), so the
    /// service restart happens silently on a background thread — zero prompts.
    /// </summary>
    private void ForceToggle()
    {
        if (_busy) return;
        _busy = true;
        _btnToggle.Enabled = false;

        bool newState = !_isMono;

        _lblStatus.Text      = Strings.StatusWait;
        _lblStatus.ForeColor = Color.FromArgb(255, 190, 50);
        _btnToggle.Text      = "…";
        _btnToggle.Invalidate();

        // !! Everything runs on background thread — UI thread never blocks !!
        System.Threading.Tasks.Task.Run(() =>
        {
            try
            {
                // Write registry in background (WM_SETTINGCHANGE can take 500 ms+)
                WriteRegistryState(newState);

                // Kill audiodg.exe first — helps audiosrv stop faster by removing
                // the child APO host before the graceful shutdown sequence.
                TryKillAudioDG();

                // Full audiosrv restart — only confirmed method to apply mono.
                RestartAudioService();

                Invoke(() =>
                {
                    _isMono = newState;
                    LoadDevices();
                    ApplyState();
                    _btnToggle.Enabled = true;
                    _busy = false;
                    ShowOverlay(newState);
                });
            }
            catch (Exception ex)
            {
                WriteRegistryState(!newState); // revert
                Invoke(() =>
                {
                    MessageBox.Show($"{Strings.ErrorPrefix}\n{ex.Message}", APP_NAME);
                    ApplyState();
                    _btnToggle.Enabled = true;
                    _busy = false;
                });
            }
        });
    }

    /// <summary>
    /// Kills audiodg.exe — the child APO host process of audiosrv.
    /// audiosrv automatically relaunches it within ~500 ms, re-reading
    /// the registry (including the mono setting) on the way back up.
    /// Much faster than a full audiosrv restart (~5 s).
    /// Returns true if at least one audiodg process was found and killed.
    /// </summary>
    private static bool TryKillAudioDG()
    {
        try
        {
            var procs = Process.GetProcessesByName("audiodg");
            if (procs.Length == 0) return false;
            foreach (var p in procs)
            {
                try { p.Kill(); } catch { /* may already be gone */ }
                finally { p.Dispose(); }
            }
            // Wait for audiosrv to relaunch the child and settle
            Thread.Sleep(700);
            return true;
        }
        catch { return false; }
    }

    /// <summary>
    /// Restarts audiosrv using ServiceController with tight 100 ms polling —
    /// faster than spawning cmd.exe with "net stop/start".
    /// </summary>
    private static void RestartAudioService()
    {
        using var sc = new System.ServiceProcess.ServiceController("audiosrv");

        if (sc.Status == System.ServiceProcess.ServiceControllerStatus.Running)
        {
            sc.Stop();
            var deadline = DateTime.UtcNow.AddSeconds(10);
            while (sc.Status != System.ServiceProcess.ServiceControllerStatus.Stopped
                   && DateTime.UtcNow < deadline)
            {
                Thread.Sleep(50);   // tighter polling — detect stop ASAP
                sc.Refresh();
            }
        }

        sc.Start();
        var startDeadline = DateTime.UtcNow.AddSeconds(10);
        while (sc.Status != System.ServiceProcess.ServiceControllerStatus.Running
               && DateTime.UtcNow < startDeadline)
        {
            Thread.Sleep(50);   // tighter polling — detect running ASAP
            sc.Refresh();
        }

        Thread.Sleep(200); // minimal settle — audio sessions reconnect quickly
    }

    // ─── Hotkey ───────────────────────────────────────────────────────────────

    private void OnSetHotkeyClick(object? sender, EventArgs e)
    {
        if (_recordingHotkey) { CancelHotkeyRecording(); return; }
        _recordingHotkey = true;
        _btnSetHotkey.Text      = Strings.HotkeyRecording;
        _btnSetHotkey.ForeColor = Color.FromArgb(255, 190, 50);
        _txtHotkey.Text         = Strings.HotkeyPrompt;
        _txtHotkey.BackColor   = Color.FromArgb(45, 42, 28);
    }

    private void OnFormKeyDown(object? sender, KeyEventArgs e)
    {
        if (!_recordingHotkey) return;

        // Ignore lone modifier keys — wait for a real key
        if (e.KeyCode is Keys.ControlKey or Keys.ShiftKey or Keys.Menu
                      or Keys.LWin or Keys.RWin or Keys.None)
            return;

        if (e.KeyCode == Keys.Escape) { CancelHotkeyRecording(); e.Handled = true; return; }

        // Require at least one modifier
        if (!e.Control && !e.Alt && !e.Shift) return;

        uint mods = 0;
        if (e.Control) mods |= MOD_CONTROL;
        if (e.Alt)     mods |= MOD_ALT;
        if (e.Shift)   mods |= MOD_SHIFT;

        SetHotkey(mods, (uint)e.KeyCode);
        e.Handled = true;
        e.SuppressKeyPress = true;
    }

    private void SetHotkey(uint modifiers, uint vk)
    {
        if (_hotkeyRegistered) { UnregisterHotKey(Handle, HOTKEY_ID); _hotkeyRegistered = false; }

        if (RegisterHotKey(Handle, HOTKEY_ID, modifiers | MOD_NOREPEAT, vk))
        {
            _hotkeyModifiers = modifiers;
            _hotkeyVk        = vk;
            _hotkeyRegistered = true;
            _txtHotkey.Text      = HotkeyToString(modifiers, vk);
            _txtHotkey.BackColor = Color.FromArgb(38, 38, 38);

            RegSet(APP_KEY, "HotkeyMods", (int)modifiers);
            RegSet(APP_KEY, "HotkeyVk",   (int)vk);
        }
        else
        {
            _txtHotkey.Text      = Strings.HotkeyInUse;
            _txtHotkey.BackColor = Color.FromArgb(60, 28, 28);
        }
        _recordingHotkey        = false;
        _btnSetHotkey.Text      = Strings.HotkeySet;
        _btnSetHotkey.ForeColor = Color.FromArgb(200, 200, 200);
    }

    private void ClearHotkey()
    {
        if (_hotkeyRegistered) { UnregisterHotKey(Handle, HOTKEY_ID); _hotkeyRegistered = false; }
        _hotkeyModifiers = 0; _hotkeyVk = 0;
        _txtHotkey.Text      = Strings.HotkeyNone;
        _txtHotkey.BackColor = Color.FromArgb(38, 38, 38);
        RegSet(APP_KEY, "HotkeyMods", 0);
        RegSet(APP_KEY, "HotkeyVk",   0);
    }

    private void CancelHotkeyRecording()
    {
        _recordingHotkey        = false;
        _btnSetHotkey.Text      = Strings.HotkeySet;
        _btnSetHotkey.ForeColor = Color.FromArgb(200, 200, 200);
        _txtHotkey.Text         = _hotkeyRegistered
            ? HotkeyToString(_hotkeyModifiers, _hotkeyVk)
            : Strings.HotkeyNone;
        _txtHotkey.BackColor    = Color.FromArgb(38, 38, 38);
    }

    private void RegisterSavedHotkey()
    {
        var mods = (uint)RegGet<int>(APP_KEY, "HotkeyMods");
        var vk   = (uint)RegGet<int>(APP_KEY, "HotkeyVk");
        if (vk == 0) return; // nothing saved
        if (RegisterHotKey(Handle, HOTKEY_ID, mods | MOD_NOREPEAT, vk))
        {
            _hotkeyModifiers  = mods;
            _hotkeyVk         = vk;
            _hotkeyRegistered = true;
            _txtHotkey.Text = HotkeyToString(mods, vk);
        }
    }

    private static string HotkeyToString(uint mods, uint vk)
    {
        var parts = new List<string>();
        if ((mods & MOD_CONTROL) != 0) parts.Add("Ctrl");
        if ((mods & MOD_ALT)     != 0) parts.Add("Alt");
        if ((mods & MOD_SHIFT)   != 0) parts.Add("Shift");
        if ((mods & MOD_WIN)     != 0) parts.Add("Win");
        var key = (Keys)vk;
        string keyName = key switch
        {
            >= Keys.A and <= Keys.Z   => ((char)('A' + (key - Keys.A))).ToString(),
            >= Keys.D0 and <= Keys.D9 => ((char)('0' + (key - Keys.D0))).ToString(),
            >= Keys.F1 and <= Keys.F24 => $"F{key - Keys.F1 + 1}",
            >= Keys.NumPad0 and <= Keys.NumPad9 => $"Num{key - Keys.NumPad0}",
            Keys.Space    => Strings.KeySpace,
            Keys.Return   => "Enter",
            Keys.Tab      => "Tab",
            Keys.Back     => "Backspace",
            Keys.Delete   => "Delete",
            Keys.Insert   => "Insert",
            Keys.Home     => "Home",
            Keys.End      => "End",
            Keys.PageUp   => "PgUp",
            Keys.PageDown => "PgDn",
            Keys.Up       => "↑",
            Keys.Down     => "↓",
            Keys.Left     => "←",
            Keys.Right    => "→",
            Keys.OemMinus => "-",
            Keys.Oemplus  => "+",
            Keys.Oemcomma => ",",
            Keys.OemPeriod => ".",
            _             => key.ToString()
        };
        parts.Add(keyName);
        return string.Join(" + ", parts);
    }

    private void ApplyState()
    {
        _btnToggle.Text      = _isMono ? Strings.BtnMono : Strings.BtnStereo;
        _btnToggle.BackColor = _isMono ? Color.FromArgb(0, 103, 192) : Color.FromArgb(58, 58, 58);
        _btnToggle.Invalidate();

        _lblStatus.Text      = _isMono ? Strings.StatusMono : Strings.StatusStereo;
        _lblStatus.ForeColor = _isMono
            ? Color.FromArgb(90, 170, 255)
            : Color.FromArgb(140, 140, 140);

        _tray.Icon           = _isMono ? _iconMono : _iconStereo;
        _tray.Text           = _isMono ? Strings.TrayMono : Strings.TrayStereo;
        _trayToggleItem.Text = _isMono ? Strings.TrayToStereo : Strings.TrayToMono;
    }

    // ─── Registry ─────────────────────────────────────────────────────────────

    private static void WriteRegistryState(bool mono)
    {
        try
        {
            using var k = Registry.CurrentUser.CreateSubKey(AUDIO_KEY);
            k.SetValue(MONO_VALUE, mono ? 1 : 0, RegistryValueKind.DWord);

            using var a = Registry.CurrentUser.CreateSubKey(ACCESS_KEY);
            var cur = a.GetValue(CONFIG_VAL) as string ?? "";
            var set = new HashSet<string>(
                cur.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries),
                StringComparer.OrdinalIgnoreCase);
            if (mono) set.Add(MONO_FEAT); else set.Remove(MONO_FEAT);
            a.SetValue(CONFIG_VAL, string.Join(",", set), RegistryValueKind.String);

            // Soft notification (best-effort)
            SendMessageTimeout(HWND_BROADCAST, WM_SETTINGCHANGE, UIntPtr.Zero,
                ACCESS_KEY, SMTO_ABORTIFHUNG | SMTO_NOTIMEOUTIFHUNG, 2000, out _);
            SendMessageTimeout(HWND_BROADCAST, WM_SETTINGCHANGE, UIntPtr.Zero,
                "AccessibilityMonoMixState", SMTO_ABORTIFHUNG, 1000, out _);

            var at = new ACCESSTIMEOUT { cbSize = (uint)Marshal.SizeOf<ACCESSTIMEOUT>() };
            if (SystemParametersInfo(SPI_GETACCESSTIMEOUT, at.cbSize, ref at, 0))
                SystemParametersInfo(SPI_SETACCESSTIMEOUT, at.cbSize, ref at,
                    SPIF_UPDATEINIFILE | SPIF_SENDCHANGE);
        }
        catch { }
    }

    private static T? RegGet<T>(string subkey, string name)
    {
        try { using var k = Registry.CurrentUser.OpenSubKey(subkey); return (T?)k?.GetValue(name); }
        catch { return default; }
    }

    private static void RegSet(string subkey, string name, object value)
    {
        try { using var k = Registry.CurrentUser.CreateSubKey(subkey); k.SetValue(name, value); }
        catch { }
    }

    private void ApplyStartup(bool enable)
    {
        try
        {
            // Remove legacy registry Run entry (migration from older builds)
            using (var k = Registry.CurrentUser.OpenSubKey(RUN_KEY, writable: true))
                k?.DeleteValue(APP_NAME, throwOnMissingValue: false);

            if (enable)
            {
                // Task Scheduler with /rl HIGHEST = runs elevated at logon, zero UAC prompt
                var tr = $"\\\"{Application.ExecutablePath}\\\" --tray";
                RunSchtasks($"/create /tn \"{APP_NAME}\" /tr \"{tr}\" /sc ONLOGON /rl HIGHEST /f");
            }
            else
            {
                RunSchtasks($"/delete /tn \"{APP_NAME}\" /f", ignoreExitCode: true);
            }
        }
        catch (Exception ex) { MessageBox.Show($"{Strings.StartupError}\n{ex.Message}", APP_NAME); }
    }

    private static bool IsStartupTaskRegistered()
    {
        try
        {
            using var p = Process.Start(new ProcessStartInfo
            {
                FileName        = "schtasks.exe",
                Arguments       = $"/query /tn \"{APP_NAME}\"",
                UseShellExecute = false,
                CreateNoWindow  = true
            })!;
            p.WaitForExit(3000);
            return p.ExitCode == 0;
        }
        catch { return false; }
    }

    private static void RunSchtasks(string args, bool ignoreExitCode = false)
    {
        using var p = Process.Start(new ProcessStartInfo
        {
            FileName               = "schtasks.exe",
            Arguments              = args,
            UseShellExecute        = false,
            CreateNoWindow         = true,
            RedirectStandardError  = true
        })!;
        var err = p.StandardError.ReadToEnd(); // read before WaitForExit to avoid deadlock
        p.WaitForExit(5000);
        if (!ignoreExitCode && p.ExitCode != 0 && !string.IsNullOrWhiteSpace(err))
            throw new InvalidOperationException(err.Trim());
    }

    // ─── Window ───────────────────────────────────────────────────────────────

    private void RestoreWindow()
    {
        Show(); WindowState = FormWindowState.Normal;
        ShowInTaskbar = true; Activate(); BringToFront();
    }

    private void ExitApp()
    {
        _exitRequested = true; _tray.Visible = false; Application.Exit();
    }

    // ─── Button painting ──────────────────────────────────────────────────────

    private void PaintButton(object? sender, PaintEventArgs e)
    {
        var btn = (Button)sender!;
        var g   = e.Graphics;
        g.SmoothingMode     = SmoothingMode.AntiAlias;
        g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

        bool hover   = btn.ClientRectangle.Contains(btn.PointToClient(Cursor.Position));
        bool pressed = hover && (MouseButtons & MouseButtons.Left) != 0;
        Color c      = btn.BackColor;
        if      (pressed) c = Adj(c, -30);
        else if (hover)   c = Adj(c, +20);

        using var fill = new SolidBrush(c);
        g.FillRectangle(fill, btn.ClientRectangle);
        using var sf = new StringFormat
            { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
        g.DrawString(btn.Text, btn.Font, Brushes.White,
            new RectangleF(0, 0, btn.Width, btn.Height), sf);
    }

    // ─── Helpers ──────────────────────────────────────────────────────────────

    private static void RoundRegion(Control c, int r)
    {
        var path = new GraphicsPath();
        int d = r * 2;
        var rect = new Rectangle(0, 0, c.Width, c.Height);
        path.AddArc(rect.X,         rect.Y,          d, d, 180, 90);
        path.AddArc(rect.Right - d, rect.Y,          d, d, 270, 90);
        path.AddArc(rect.Right - d, rect.Bottom - d, d, d,   0, 90);
        path.AddArc(rect.X,         rect.Bottom - d, d, d,  90, 90);
        path.CloseFigure();
        c.Region = new Region(path);
    }

    private static Color Adj(Color c, int d) =>
        Color.FromArgb(c.A, Math.Clamp(c.R + d, 0, 255),
            Math.Clamp(c.G + d, 0, 255), Math.Clamp(c.B + d, 0, 255));

    private static Icon MakeIcon(Color bg, string letter)
    {
        using var bmp = new Bitmap(32, 32);
        using var g   = Graphics.FromImage(bmp);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.Clear(Color.Transparent);
        g.FillEllipse(new SolidBrush(bg), 1, 1, 30, 30);
        using var font = new Font("Segoe UI", 15f, FontStyle.Bold, GraphicsUnit.Pixel);
        using var sf   = new StringFormat
            { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
        g.DrawString(letter, font, Brushes.White, new RectangleF(0, 0, 32, 32), sf);
        var h = bmp.GetHicon();
        var i = (Icon)Icon.FromHandle(h).Clone();
        DestroyIcon(h);
        return i;
    }

    private static Label MkLabel(string text, Font font, Color color, Point loc) => new()
    {
        Text = text, Font = font, ForeColor = color,
        AutoSize = true, Location = loc, BackColor = Color.Transparent
    };

    private static CheckBox MkCheckBox(string text, Point loc, bool chk) => new()
    {
        Text = text, AutoSize = true, Location = loc, Checked = chk,
        ForeColor = Color.FromArgb(185, 185, 185), FlatStyle = FlatStyle.Flat,
        BackColor = Color.Transparent
    };

    private void ShowOverlay(bool isMono)
    {
        _overlay?.Close();
        _overlay = new OverlayForm(isMono);
        _overlay.FormClosed += (_, _) => _overlay = null;
        _overlay.Show();
    }
}

// ─── Dark menu renderer ───────────────────────────────────────────────────────

internal sealed class DarkMenuRenderer : ToolStripProfessionalRenderer
{
    private static readonly Color Bg    = Color.FromArgb(38, 38, 38);
    private static readonly Color Hover = Color.FromArgb(60, 60, 60);
    private static readonly Color Line  = Color.FromArgb(62, 62, 62);

    public DarkMenuRenderer() : base(new DarkColorTable()) { }

    protected override void OnRenderToolStripBackground(ToolStripRenderEventArgs e)
    { using var b = new SolidBrush(Bg); e.Graphics.FillRectangle(b, e.AffectedBounds); }

    protected override void OnRenderMenuItemBackground(ToolStripItemRenderEventArgs e)
    { using var b = new SolidBrush(e.Item.Selected ? Hover : Bg);
      e.Graphics.FillRectangle(b, e.Item.ContentRectangle); }

    protected override void OnRenderItemText(ToolStripItemTextRenderEventArgs e)
    { e.TextColor = e.Item.Enabled ? Color.FromArgb(218, 218, 218) : Color.FromArgb(90, 90, 90);
      base.OnRenderItemText(e); }

    protected override void OnRenderSeparator(ToolStripSeparatorRenderEventArgs e)
    { int y = e.Item.ContentRectangle.Height / 2;
      using var p = new Pen(Line);
      e.Graphics.DrawLine(p, e.Item.ContentRectangle.Left, y, e.Item.ContentRectangle.Right, y); }
}

internal sealed class DarkColorTable : ProfessionalColorTable
{
    private static readonly Color C = Color.FromArgb(38, 38, 38);
    private static readonly Color B = Color.FromArgb(62, 62, 62);
    public override Color MenuBorder                  => B;
    public override Color ToolStripDropDownBackground => C;
    public override Color ImageMarginGradientBegin    => C;
    public override Color ImageMarginGradientMiddle   => C;
    public override Color ImageMarginGradientEnd      => C;
}

// ─── Audio state overlay ──────────────────────────────────────────────────────

internal sealed class OverlayForm : Form
{
    private readonly bool   _isMono;
    private readonly System.Windows.Forms.Timer _timer;
    private int _tick;

    // Timing (each tick = 20 ms)
    private const int FadeInTicks  =  8;   //  160 ms fade-in
    private const int HoldTicks    = 135;  // 2700 ms hold
    private const int FadeOutTicks = 12;   //  240 ms fade-out
    private const int TotalTicks   = FadeInTicks + HoldTicks + FadeOutTicks;

    private const int W = 230;
    private const int H = 76;

    public OverlayForm(bool isMono)
    {
        _isMono = isMono;

        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar   = false;
        TopMost         = true;
        StartPosition   = FormStartPosition.Manual;
        BackColor       = Color.Black;
        TransparencyKey = Color.Black;
        Size            = new Size(W, H);
        Opacity         = 0;

        // Bottom-right of the screen where the mouse cursor currently is
        var wa = Screen.FromPoint(Cursor.Position).WorkingArea;
        Location = new Point(wa.Right - W - 24, wa.Bottom - H - 24);

        SetStyle(ControlStyles.OptimizedDoubleBuffer |
                 ControlStyles.AllPaintingInWmPaint  |
                 ControlStyles.UserPaint, true);

        _timer = new System.Windows.Forms.Timer { Interval = 20 };
        _timer.Tick += OnTick;
        _timer.Start();
    }

    // Prevent stealing focus
    protected override bool ShowWithoutActivation => true;
    protected override CreateParams CreateParams
    {
        get
        {
            var cp   = base.CreateParams;
            cp.ExStyle |= 0x08000000; // WS_EX_NOACTIVATE
            cp.ExStyle |= 0x00000080; // WS_EX_TOOLWINDOW
            return cp;
        }
    }

    private void OnTick(object? sender, EventArgs e)
    {
        _tick++;
        Opacity = _tick switch
        {
            <= FadeInTicks                          => (double)_tick / FadeInTicks,
            <= FadeInTicks + HoldTicks              => 1.0,
            <= TotalTicks                           => 1.0 - (double)(_tick - FadeInTicks - HoldTicks) / FadeOutTicks,
            _                                       => 0
        };
        if (_tick > TotalTicks) { _timer.Stop(); Close(); }
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        var g = e.Graphics;
        g.SmoothingMode     = SmoothingMode.AntiAlias;
        g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

        // ── Background rounded rect ───────────────────────────────────────────
        const int r = 12;
        using var path = BuildRoundedPath(0, 0, W, H, r);
        using (var bg = new SolidBrush(Color.FromArgb(30, 30, 30)))
            g.FillPath(bg, path);
        using (var border = new Pen(Color.FromArgb(65, 65, 65), 1f))
            g.DrawPath(border, path);

        // ── Circle icon ───────────────────────────────────────────────────────
        var circleColor = _isMono ? Color.FromArgb(0, 103, 192) : Color.FromArgb(72, 72, 72);
        const int cx = 14, cy = 14, cs = 48;
        using (var circleBrush = new SolidBrush(circleColor))
            g.FillEllipse(circleBrush, cx, cy, cs, cs);

        string letter = _isMono ? "M" : "S";
        using var lf = new Font("Segoe UI", 22f, FontStyle.Bold, GraphicsUnit.Pixel);
        using var sf = new StringFormat
            { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
        g.DrawString(letter, lf, Brushes.White, new RectangleF(cx, cy, cs, cs), sf);

        // ── State text ────────────────────────────────────────────────────────
        const int tx = 74;
        string mainText = _isMono ? Strings.BtnMono : Strings.BtnStereo;
        var mainColor = _isMono ? Color.FromArgb(100, 180, 255) : Color.FromArgb(230, 230, 230);
        using var mf        = new Font("Segoe UI", 20f, FontStyle.Bold, GraphicsUnit.Pixel);
        using var mainBrush = new SolidBrush(mainColor);
        g.DrawString(mainText, mf, mainBrush, new PointF(tx, 10));

        string subText = _isMono ? Strings.OverlayMonoSub : Strings.OverlayStereoSub;
        using var sf2     = new Font("Segoe UI", 11f, GraphicsUnit.Pixel);
        using var subBrush = new SolidBrush(Color.FromArgb(120, 120, 120));
        g.DrawString(subText, sf2, subBrush, new PointF(tx + 1, 38));
    }

    private static GraphicsPath BuildRoundedPath(int x, int y, int w, int h, int r)
    {
        int d = r * 2;
        var p = new GraphicsPath();
        p.AddArc(x,         y,         d, d, 180, 90);
        p.AddArc(x + w - d, y,         d, d, 270, 90);
        p.AddArc(x + w - d, y + h - d, d, d,   0, 90);
        p.AddArc(x,         y + h - d, d, d,  90, 90);
        p.CloseFigure();
        return p;
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing) _timer.Dispose();
        base.Dispose(disposing);
    }
}

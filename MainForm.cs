using System.Diagnostics;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
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
    private const int  HOTKEY_ID    = 0x4D53;

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

    // ─── Design tokens ────────────────────────────────────────────────────────

    internal static readonly Color CBg       = Color.FromArgb(12,  12,  16);
    internal static readonly Color CCard     = Color.FromArgb(22,  22,  30);
    internal static readonly Color CBorder   = Color.FromArgb(50,  50,  70);
    internal static readonly Color CAccentB  = Color.FromArgb(74, 158, 255);   // MONO
    internal static readonly Color CAccentG  = Color.FromArgb(45, 210, 140);   // STEREO
    internal static readonly Color CText     = Color.FromArgb(215, 215, 230);
    internal static readonly Color CMuted    = Color.FromArgb(90,  90, 110);
    internal static readonly Color CDim      = Color.FromArgb(48,  48,  62);

    // ─── State ────────────────────────────────────────────────────────────────

    private bool _isMono;
    private bool _busy;
    private readonly bool _startHidden;
    private bool _initialShowDone;
    private bool _exitRequested;

    private bool _recordingHotkey;
    private uint _hotkeyModifiers;
    private uint _hotkeyVk;
    private bool _hotkeyRegistered;

    private OverlayForm? _overlay;

    // ─── Controls ─────────────────────────────────────────────────────────────

    private readonly NotifyIcon        _tray;
    private readonly ContextMenuStrip  _trayMenu;
    private readonly ToolStripMenuItem _trayToggleItem;
    private readonly Button            _btnToggle;
    private readonly Label             _lblWait;
    private readonly ListBox           _listDevices;
    private readonly ToggleSwitch      _chkStartup;
    private readonly ToggleSwitch      _chkTray;
    private readonly HotkeyBox        _hotkeyBox;
    private readonly Icon              _iconMono;
    private readonly Icon              _iconStereo;
    private readonly Icon              _iconApp;

    private List<AudioDevice> _devices = new();

    // ─── Constructor ──────────────────────────────────────────────────────────

    public MainForm(bool startHidden)
    {
        _isMono      = RegGet<int>(AUDIO_KEY, MONO_VALUE) == 1;
        _startHidden = startHidden || RegGet<int>(APP_KEY, "StartInTray") == 1;

        _iconMono   = MakeIcon(CAccentB, "M");
        _iconStereo = MakeIcon(CAccentG, "S");
        _iconApp    = MakeAppIcon();
        Icon        = _iconApp;

        // ── Form ──────────────────────────────────────────────────────────────
        Text            = APP_NAME;
        Size            = new Size(400, 580);
        MinimumSize     = new Size(340, 500);
        FormBorderStyle = FormBorderStyle.Sizable;
        MaximizeBox     = true;
        StartPosition   = FormStartPosition.CenterScreen;
        BackColor       = CBg;
        ForeColor       = CText;
        Font            = new Font("Segoe UI", 10f);

        // ── Header bar ────────────────────────────────────────────────────────
        var header = new HeaderBar { Dock = DockStyle.Top, Height = 54 };
        Controls.Add(header);

        // ── Toggle button ─────────────────────────────────────────────────────
        _btnToggle = new Button
        {
            Location  = new Point(16, 62),
            Size      = new Size(ClientSize.Width - 32, 100),
            Anchor    = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            FlatStyle = FlatStyle.Flat,
            Cursor    = Cursors.Hand,
            UseVisualStyleBackColor = false,
            BackColor = Color.Transparent,
            TabStop   = false,
        };
        _btnToggle.FlatAppearance.BorderSize         = 0;
        _btnToggle.FlatAppearance.MouseOverBackColor = Color.Transparent;
        _btnToggle.FlatAppearance.MouseDownBackColor = Color.Transparent;
        _btnToggle.Paint  += PaintToggleButton;
        _btnToggle.Click  += (_, _) => ForceToggle();
        Controls.Add(_btnToggle);

        // ── Wait label (shown during toggle) ─────────────────────────────────
        _lblWait = new Label
        {
            Text      = Strings.StatusWait,
            Font      = new Font("Segoe UI", 9f),
            ForeColor = Color.FromArgb(255, 190, 50),
            BackColor = Color.Transparent,
            AutoSize  = true,
            Visible   = false,
            Anchor    = AnchorStyles.Top | AnchorStyles.Left,
        };
        Controls.Add(_lblWait);
        _lblWait.Location = new Point(18, 170);

        // ── Devices section ───────────────────────────────────────────────────
        var sectionDevices = new SectionLabel(Strings.DevicesLabel)
        {
            Location = new Point(16, 180),
            Size     = new Size(ClientSize.Width - 32, 20),
            Anchor   = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
        };
        Controls.Add(sectionDevices);

        var devCard = new CardPanel
        {
            Location = new Point(16, 204),
            Size     = new Size(ClientSize.Width - 32, ClientSize.Height - 404),
            Anchor   = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
        };
        Controls.Add(devCard);

        _listDevices = new ListBox
        {
            Location      = new Point(1, 1),
            Size          = new Size(devCard.Width - 2, devCard.Height - 2),
            Anchor        = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
            BackColor     = Color.FromArgb(18, 18, 26),
            ForeColor     = CText,
            BorderStyle   = BorderStyle.None,
            Font          = new Font("Segoe UI", 9f),
            SelectionMode = SelectionMode.None,
            DrawMode      = DrawMode.OwnerDrawFixed,
            ItemHeight    = 28,
        };
        _listDevices.DrawItem += DrawDeviceItem;
        devCard.Controls.Add(_listDevices);

        // ── Device hint ───────────────────────────────────────────────────────
        var lblHint = new Label
        {
            Text      = Strings.DeviceHint,
            Font      = new Font("Segoe UI", 7.5f),
            ForeColor = Color.FromArgb(52, 52, 66),
            BackColor = Color.Transparent,
            AutoSize  = true,
            Location  = new Point(18, ClientSize.Height - 214),
            Anchor    = AnchorStyles.Bottom | AnchorStyles.Left,
        };
        Controls.Add(lblHint);

        // ── Settings card ─────────────────────────────────────────────────────
        var settingsCard = new CardPanel
        {
            Location = new Point(16, ClientSize.Height - 198),
            Size     = new Size(ClientSize.Width - 32, 170),
            Anchor   = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
        };
        Controls.Add(settingsCard);

        // Startup toggle row
        var rowStartup = new ToggleRow(Strings.ChkStartup, IsStartupTaskRegistered());
        rowStartup.Location = new Point(0, 0);
        rowStartup.Size     = new Size(settingsCard.Width, 46);
        rowStartup.Anchor   = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        settingsCard.Controls.Add(rowStartup);
        _chkStartup = rowStartup.Toggle;
        _chkStartup.CheckedChanged += (_, _) => ApplyStartup(_chkStartup.Checked);

        // Thin divider
        var div1 = new DividerLine { Location = new Point(16, 46), Size = new Size(settingsCard.Width - 32, 1) };
        div1.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        settingsCard.Controls.Add(div1);

        // Tray toggle row
        var rowTray = new ToggleRow(Strings.ChkTray, RegGet<int>(APP_KEY, "StartInTray") == 1);
        rowTray.Location = new Point(0, 47);
        rowTray.Size     = new Size(settingsCard.Width, 46);
        rowTray.Anchor   = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        settingsCard.Controls.Add(rowTray);
        _chkTray = rowTray.Toggle;
        _chkTray.CheckedChanged += (_, _) => RegSet(APP_KEY, "StartInTray", _chkTray.Checked ? 1 : 0);

        // Thin divider
        var div2 = new DividerLine { Location = new Point(16, 93), Size = new Size(settingsCard.Width - 32, 1) };
        div2.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        settingsCard.Controls.Add(div2);

        // Hotkey row
        _hotkeyBox = new HotkeyBox
        {
            Location = new Point(0, 94),
            Size     = new Size(settingsCard.Width, 46),
            Anchor   = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
        };
        _hotkeyBox.SetClicked   += OnSetHotkeyClick;
        _hotkeyBox.ClearClicked += (_, _) => ClearHotkey();
        settingsCard.Controls.Add(_hotkeyBox);

        // ── Footer ────────────────────────────────────────────────────────────
        var footer = new Label
        {
            Text      = Strings.Footer,
            Font      = new Font("Segoe UI", 7f),
            ForeColor = Color.FromArgb(42, 42, 55),
            BackColor = Color.Transparent,
            AutoSize  = false,
            Size      = new Size(ClientSize.Width - 36, 16),
            Location  = new Point(18, ClientSize.Height - 26),
            Anchor    = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
        };
        Controls.Add(footer);

        // ── Key capture ───────────────────────────────────────────────────────
        KeyPreview = true;
        KeyDown   += OnFormKeyDown;

        // ── Tray ──────────────────────────────────────────────────────────────
        _trayToggleItem = new ToolStripMenuItem { Font = new Font("Segoe UI", 9f, FontStyle.Bold) };
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
            _iconMono.Dispose(); _iconStereo.Dispose(); _iconApp.Dispose();
        }
        base.Dispose(disposing);
    }

    // ─── Devices ──────────────────────────────────────────────────────────────

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
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

        using (var bg = new SolidBrush(Color.FromArgb(e.Index % 2 == 0 ? 18 : 21, 18, 26)))
            g.FillRectangle(bg, rect);

        // Accent dot
        var dotColor = dev.IsDefault ? CAccentG : CDim;
        using (var dot = new SolidBrush(dotColor))
            g.FillEllipse(dot, rect.X + 12, rect.Y + 10, 8, 8);

        var nameColor = dev.IsDefault ? Color.White : CMuted;
        using var nf = dev.IsDefault
            ? new Font("Segoe UI", 9f, FontStyle.Bold)
            : new Font("Segoe UI", 9f);
        using (var nb = new SolidBrush(nameColor))
            g.DrawString(dev.Name, nf, nb, new PointF(rect.X + 28, rect.Y + 5));

        if (dev.IsDefault)
        {
            float tw = g.MeasureString(dev.Name, nf).Width;
            using var tf = new Font("Segoe UI", 7.5f);
            using var tb = new SolidBrush(CAccentG);
            g.DrawString(Strings.DeviceDefault, tf, tb, new PointF(rect.X + 28 + tw + 4, rect.Y + 7));
        }

        using var div = new Pen(Color.FromArgb(28, 28, 38));
        g.DrawLine(div, rect.Left, rect.Bottom - 1, rect.Right, rect.Bottom - 1);
    }

    // ─── Toggle ───────────────────────────────────────────────────────────────

    private void ForceToggle()
    {
        if (_busy) return;
        _busy = true;
        _btnToggle.Enabled = false;
        _lblWait.Visible   = true;
        _btnToggle.Invalidate();

        bool newState = !_isMono;

        System.Threading.Tasks.Task.Run(() =>
        {
            try
            {
                WriteRegistryState(newState);
                TryKillAudioDG();
                RestartAudioService();

                Invoke(() =>
                {
                    _isMono = newState;
                    LoadDevices();
                    ApplyState();
                    _btnToggle.Enabled = true;
                    _lblWait.Visible   = false;
                    _busy = false;
                    ShowOverlay(newState);
                });
            }
            catch (Exception ex)
            {
                WriteRegistryState(!newState);
                Invoke(() =>
                {
                    MessageBox.Show($"{Strings.ErrorPrefix}\n{ex.Message}", APP_NAME);
                    ApplyState();
                    _btnToggle.Enabled = true;
                    _lblWait.Visible   = false;
                    _busy = false;
                });
            }
        });
    }

    private static bool TryKillAudioDG()
    {
        try
        {
            var procs = Process.GetProcessesByName("audiodg");
            if (procs.Length == 0) return false;
            foreach (var p in procs)
            {
                try { p.Kill(); } catch { }
                finally { p.Dispose(); }
            }
            Thread.Sleep(700);
            return true;
        }
        catch { return false; }
    }

    private static void RestartAudioService()
    {
        using var sc = new System.ServiceProcess.ServiceController("audiosrv");
        if (sc.Status == System.ServiceProcess.ServiceControllerStatus.Running)
        {
            sc.Stop();
            var d = DateTime.UtcNow.AddSeconds(10);
            while (sc.Status != System.ServiceProcess.ServiceControllerStatus.Stopped && DateTime.UtcNow < d)
            { Thread.Sleep(50); sc.Refresh(); }
        }
        sc.Start();
        var d2 = DateTime.UtcNow.AddSeconds(10);
        while (sc.Status != System.ServiceProcess.ServiceControllerStatus.Running && DateTime.UtcNow < d2)
        { Thread.Sleep(50); sc.Refresh(); }
        Thread.Sleep(200);
    }

    // ─── Hotkey ───────────────────────────────────────────────────────────────

    private void OnSetHotkeyClick(object? sender, EventArgs e)
    {
        if (_recordingHotkey) { CancelHotkeyRecording(); return; }
        _recordingHotkey = true;
        _hotkeyBox.StartRecording();
    }

    private void OnFormKeyDown(object? sender, KeyEventArgs e)
    {
        if (!_recordingHotkey) return;
        if (e.KeyCode is Keys.ControlKey or Keys.ShiftKey or Keys.Menu or Keys.LWin or Keys.RWin or Keys.None)
            return;
        if (e.KeyCode == Keys.Escape) { CancelHotkeyRecording(); e.Handled = true; return; }
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
            _hotkeyModifiers  = modifiers;
            _hotkeyVk         = vk;
            _hotkeyRegistered = true;
            _hotkeyBox.SetValue(HotkeyToString(modifiers, vk), success: true);
            RegSet(APP_KEY, "HotkeyMods", (int)modifiers);
            RegSet(APP_KEY, "HotkeyVk",   (int)vk);
        }
        else
        {
            _hotkeyBox.SetValue(Strings.HotkeyInUse, success: false);
        }
        _recordingHotkey = false;
    }

    private void ClearHotkey()
    {
        if (_hotkeyRegistered) { UnregisterHotKey(Handle, HOTKEY_ID); _hotkeyRegistered = false; }
        _hotkeyModifiers = 0; _hotkeyVk = 0;
        _hotkeyBox.SetValue(Strings.HotkeyNone, success: true);
        RegSet(APP_KEY, "HotkeyMods", 0);
        RegSet(APP_KEY, "HotkeyVk",   0);
    }

    private void CancelHotkeyRecording()
    {
        _recordingHotkey = false;
        _hotkeyBox.CancelRecording(_hotkeyRegistered
            ? HotkeyToString(_hotkeyModifiers, _hotkeyVk)
            : Strings.HotkeyNone);
    }

    private void RegisterSavedHotkey()
    {
        var mods = (uint)RegGet<int>(APP_KEY, "HotkeyMods");
        var vk   = (uint)RegGet<int>(APP_KEY, "HotkeyVk");
        if (vk == 0) return;
        if (RegisterHotKey(Handle, HOTKEY_ID, mods | MOD_NOREPEAT, vk))
        {
            _hotkeyModifiers  = mods;
            _hotkeyVk         = vk;
            _hotkeyRegistered = true;
            _hotkeyBox.SetValue(HotkeyToString(mods, vk), success: true);
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
            >= Keys.A and <= Keys.Z    => ((char)('A' + (key - Keys.A))).ToString(),
            >= Keys.D0 and <= Keys.D9  => ((char)('0' + (key - Keys.D0))).ToString(),
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
            Keys.Up => "↑", Keys.Down => "↓", Keys.Left => "←", Keys.Right => "→",
            Keys.OemMinus => "-", Keys.Oemplus => "+",
            Keys.Oemcomma => ",", Keys.OemPeriod => ".",
            _ => key.ToString()
        };
        parts.Add(keyName);
        return string.Join(" + ", parts);
    }

    // ─── Apply state ──────────────────────────────────────────────────────────

    private void ApplyState()
    {
        _btnToggle.Invalidate();
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
            using (var k = Registry.CurrentUser.OpenSubKey(RUN_KEY, writable: true))
                k?.DeleteValue(APP_NAME, throwOnMissingValue: false);
            if (enable)
            {
                var tr = $"\\\"{Application.ExecutablePath}\\\" --tray";
                RunSchtasks($"/create /tn \"{APP_NAME}\" /tr \"{tr}\" /sc ONLOGON /rl HIGHEST /f");
            }
            else RunSchtasks($"/delete /tn \"{APP_NAME}\" /f", ignoreExitCode: true);
        }
        catch (Exception ex) { MessageBox.Show($"{Strings.StartupError}\n{ex.Message}", APP_NAME); }
    }

    private static bool IsStartupTaskRegistered()
    {
        try
        {
            using var p = Process.Start(new ProcessStartInfo
            {
                FileName = "schtasks.exe", Arguments = $"/query /tn \"{APP_NAME}\"",
                UseShellExecute = false, CreateNoWindow = true
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
            FileName = "schtasks.exe", Arguments = args,
            UseShellExecute = false, CreateNoWindow = true, RedirectStandardError = true
        })!;
        var err = p.StandardError.ReadToEnd();
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

    // ─── Toggle button paint ──────────────────────────────────────────────────

    private void PaintToggleButton(object? sender, PaintEventArgs e)
    {
        var btn = (Button)sender!;
        var g   = e.Graphics;
        g.SmoothingMode     = SmoothingMode.AntiAlias;
        g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

        bool hover   = btn.ClientRectangle.Contains(btn.PointToClient(Cursor.Position));
        bool pressed = hover && (MouseButtons & MouseButtons.Left) != 0;

        Color accent = _isMono ? CAccentB : CAccentG;

        // Paint form background colour in the full rect first, so rounded
        // corners of the card below appear to "cut into" the form surface.
        using (var bgBrush = new SolidBrush(CBg))
            g.FillRectangle(bgBrush, btn.ClientRectangle);

        // Card fill
        Color fill = _isMono
            ? Color.FromArgb(pressed ? 8 : hover ? 16 : 12, 20, 40)
            : Color.FromArgb(pressed ? 8 : hover ? 16 : 10, 28, 24);

        const int r = 14;
        using var path = RoundedPath(0, 0, btn.Width - 1, btn.Height - 1, r);
        using (var fb = new SolidBrush(fill)) g.FillPath(fb, path);

        // Border — glow on hover
        float bw = hover ? 1.8f : 1.2f;
        var   bc = hover ? Color.FromArgb(Math.Min(accent.R + 30, 255), Math.Min(accent.G + 30, 255), Math.Min(accent.B + 30, 255))
                         : Color.FromArgb((int)(accent.R * 0.7f), (int)(accent.G * 0.7f), (int)(accent.B * 0.7f));
        using (var pen = new Pen(bc, bw)) g.DrawPath(pen, path);

        // ── Left badge circle ─────────────────────────────────────────────────
        int cs  = btn.Height - 40;
        int cx  = 22;
        int cty = (btn.Height - cs) / 2;

        // Filled badge background
        var badgeBg = Color.FromArgb(30, accent.R, accent.G, accent.B);
        using (var bb = new SolidBrush(badgeBg)) g.FillEllipse(bb, cx, cty, cs, cs);
        using (var bp = new Pen(accent, 1.5f)) g.DrawEllipse(bp, cx, cty, cs, cs);

        string letter = _isMono ? "M" : "S";
        using var lf  = new Font("Segoe UI", cs * 0.50f, FontStyle.Bold, GraphicsUnit.Pixel);
        using var lsf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
        using var lb  = new SolidBrush(accent);
        g.DrawString(letter, lf, lb, new RectangleF(cx, cty, cs, cs), lsf);

        // ── Mode label ────────────────────────────────────────────────────────
        int tx  = cx + cs + 20;
        int mid = btn.Height / 2;

        using var mf = new Font("Segoe UI", 24f, FontStyle.Bold, GraphicsUnit.Pixel);
        using var mb = new SolidBrush(accent);
        g.DrawString(_isMono ? Strings.BtnMono : Strings.BtnStereo, mf, mb, new PointF(tx, mid - 17));

        // ── Subtitle ──────────────────────────────────────────────────────────
        string sub = _isMono ? Strings.StatusMono : Strings.StatusStereo;
        using var sf = new Font("Segoe UI", 10f, GraphicsUnit.Pixel);
        using var sb = new SolidBrush(CMuted);
        g.DrawString(sub, sf, sb, new PointF(tx + 1, mid + 8));

        // ── Arrow hint (right edge) ────────────────────────────────────────────
        string arrow = "›";
        using var af  = new Font("Segoe UI", 18f, FontStyle.Regular, GraphicsUnit.Pixel);
        using var asf = new StringFormat { Alignment = StringAlignment.Far, LineAlignment = StringAlignment.Center };
        using var arb = new SolidBrush(Color.FromArgb(hover ? 80 : 40, accent.R, accent.G, accent.B));
        g.DrawString(arrow, af, arb, new RectangleF(0, 0, btn.Width - 14, btn.Height), asf);
    }

    // ─── Shared helpers ───────────────────────────────────────────────────────

    internal static GraphicsPath RoundedPath(int x, int y, int w, int h, int r)
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

    internal static Region RoundedRegion(int w, int h, int r)
    {
        using var p = RoundedPath(0, 0, w - 1, h - 1, r);
        return new Region(p);
    }

    private static Icon MakeIcon(Color accent, string letter)
    {
        using var bmp = new Bitmap(32, 32);
        using var g   = Graphics.FromImage(bmp);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.Clear(Color.Transparent);
        using var pen  = new Pen(accent, 2f);
        g.DrawEllipse(pen, 2, 2, 27, 27);
        using var font  = new Font("Segoe UI", 15f, FontStyle.Bold, GraphicsUnit.Pixel);
        using var sf    = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
        using var brush = new SolidBrush(accent);
        g.DrawString(letter, font, brush, new RectangleF(0, 0, 32, 32), sf);
        var h = bmp.GetHicon();
        var i = (Icon)Icon.FromHandle(h).Clone();
        DestroyIcon(h);
        return i;
    }

    /// <summary>
    /// Generates the main application icon: golden speaker on dark navy circle.
    /// </summary>
    private static Icon MakeAppIcon()
    {
        using var bmp = new Bitmap(64, 64);
        using var g   = Graphics.FromImage(bmp);
        g.SmoothingMode     = SmoothingMode.AntiAlias;
        g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
        g.Clear(Color.Transparent);

        // ── Background circle ─────────────────────────────────────────────────
        using var bgBrush = new SolidBrush(Color.FromArgb(18, 24, 52));
        g.FillEllipse(bgBrush, 0, 0, 63, 63);
        using var bgPen = new Pen(Color.FromArgb(50, 60, 100), 1.5f);
        g.DrawEllipse(bgPen, 1, 1, 61, 61);

        // Gold colour
        var gold      = Color.FromArgb(220, 168, 36);
        var goldLight = Color.FromArgb(255, 210, 80);
        var goldDim   = Color.FromArgb(160, 120, 20);

        // ── Speaker body (rectangle, left-center) ─────────────────────────────
        float bx = 12f, by = 23f, bw = 10f, bh = 18f;
        using var bodyBrush = new LinearGradientBrush(
            new PointF(bx, 0), new PointF(bx + bw, 0), goldDim, goldLight);
        g.FillRectangle(bodyBrush, bx, by, bw, bh);

        // ── Speaker cone (trapezoid expanding right) ──────────────────────────
        var cone = new PointF[]
        {
            new(bx + bw, by),
            new(bx + bw, by + bh),
            new(bx + bw + 16f, by + bh + 8f),
            new(bx + bw + 16f, by - 8f),
        };
        using var coneBrush = new LinearGradientBrush(
            new PointF(bx + bw, 0), new PointF(bx + bw + 16f, 0), goldLight, gold);
        g.FillPolygon(coneBrush, cone);

        // Cone outline
        using var outlinePen = new Pen(goldDim, 0.8f);
        g.DrawPolygon(outlinePen, cone);
        g.DrawRectangle(outlinePen, bx, by, bw, bh);

        // ── Sound waves (arcs, right of cone) ─────────────────────────────────
        float wx = bx + bw + 18f;
        float cy = 32f;

        using var wave1 = new Pen(gold, 2f) { StartCap = LineCap.Round, EndCap = LineCap.Round };
        using var wave2 = new Pen(Color.FromArgb(180, gold.R, gold.G, gold.B), 1.8f)
            { StartCap = LineCap.Round, EndCap = LineCap.Round };
        using var wave3 = new Pen(Color.FromArgb(110, gold.R, gold.G, gold.B), 1.5f)
            { StartCap = LineCap.Round, EndCap = LineCap.Round };

        g.DrawArc(wave1, wx,       cy - 7f,  6f,  14f, -55, 110);
        g.DrawArc(wave2, wx + 5f,  cy - 11f, 8f,  22f, -55, 110);
        g.DrawArc(wave3, wx + 11f, cy - 15f, 9f,  30f, -55, 110);

        // Convert to icon
        var hicon = bmp.GetHicon();
        var icon  = (Icon)Icon.FromHandle(hicon).Clone();
        DestroyIcon(hicon);
        return icon;
    }

    private void ShowOverlay(bool isMono)
    {
        _overlay?.Close();
        _overlay = new OverlayForm(isMono);
        _overlay.FormClosed += (_, _) => _overlay = null;
        _overlay.Show();
    }
}

// ─── HeaderBar ────────────────────────────────────────────────────────────────

internal sealed class HeaderBar : Control
{
    public HeaderBar()
    {
        SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                 ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw |
                 ControlStyles.SupportsTransparentBackColor, true);
        BackColor = Color.Transparent;
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        var g = e.Graphics;
        g.SmoothingMode     = SmoothingMode.AntiAlias;
        g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

        // Bottom border line
        using var pen = new Pen(MainForm.CBorder, 1f);
        g.DrawLine(pen, 0, Height - 1, Width, Height - 1);

        // App title
        using var tf = new Font("Segoe UI", 11f, FontStyle.Bold);
        using var tb = new SolidBrush(MainForm.CAccentB);
        g.DrawString(Strings.AppTitle, tf, tb, new PointF(18, 16));

        // Version tag pill
        const string ver = "v1.0";
        using var vf   = new Font("Segoe UI", 7.5f, FontStyle.Bold);
        float vw = g.MeasureString(ver, vf).Width + 10;
        float vx = 18 + g.MeasureString(Strings.AppTitle, tf).Width + 8;
        float vy = 18;
        using var pillPath = MainForm.RoundedPath((int)vx, (int)vy, (int)vw, 16, 4);
        using (var pb = new SolidBrush(Color.FromArgb(40, MainForm.CAccentB.R, MainForm.CAccentB.G, MainForm.CAccentB.B)))
            g.FillPath(pb, pillPath);
        using (var pp = new Pen(Color.FromArgb(80, MainForm.CAccentB.R, MainForm.CAccentB.G, MainForm.CAccentB.B), 1f))
            g.DrawPath(pp, pillPath);
        using var vb = new SolidBrush(MainForm.CAccentB);
        using var vsf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
        g.DrawString(ver, vf, vb, new RectangleF(vx, vy, vw, 16), vsf);
    }
}

// ─── SectionLabel ─────────────────────────────────────────────────────────────

internal sealed class SectionLabel : Control
{
    private readonly string _text;

    public SectionLabel(string text)
    {
        _text = text.ToUpperInvariant();
        SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                 ControlStyles.OptimizedDoubleBuffer | ControlStyles.SupportsTransparentBackColor, true);
        BackColor = Color.Transparent;
        Height    = 20;
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        var g = e.Graphics;
        g.SmoothingMode     = SmoothingMode.AntiAlias;
        g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

        using var f  = new Font("Segoe UI", 7.5f, FontStyle.Bold);
        using var b  = new SolidBrush(Color.FromArgb(110, MainForm.CAccentB.R, MainForm.CAccentB.G, MainForm.CAccentB.B));
        using var lb = new SolidBrush(MainForm.CDim);

        float tw = g.MeasureString(_text, f).Width;
        g.DrawString(_text, f, b, new PointF(0, 2));

        // Trailing line
        int lx = (int)tw + 8;
        int ly = Height / 2;
        using var lp = new Pen(MainForm.CDim, 1f);
        g.DrawLine(lp, lx, ly, Width, ly);
    }
}

// ─── CardPanel ────────────────────────────────────────────────────────────────

internal sealed class CardPanel : Panel
{
    public CardPanel()
    {
        SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                 ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        var g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;
        using var path = MainForm.RoundedPath(0, 0, Width - 1, Height - 1, 10);
        using (var fb = new SolidBrush(MainForm.CCard)) g.FillPath(fb, path);
        using (var bp = new Pen(MainForm.CBorder, 1f))  g.DrawPath(bp, path);
    }

    protected override void OnPaintBackground(PaintEventArgs e) { }
}

// ─── DividerLine ──────────────────────────────────────────────────────────────

internal sealed class DividerLine : Control
{
    public DividerLine()
    {
        SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                 ControlStyles.SupportsTransparentBackColor, true);
        BackColor = Color.Transparent;
        Height = 1;
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        using var p = new Pen(Color.FromArgb(38, 38, 55), 1f);
        e.Graphics.DrawLine(p, 0, 0, Width, 0);
    }
}

// ─── ToggleSwitch ─────────────────────────────────────────────────────────────

internal sealed class ToggleSwitch : Control
{
    private bool _checked;
    private bool _hover;
    private const int TW = 36, TH = 20;

    public bool Checked
    {
        get => _checked;
        set { _checked = value; Invalidate(); }
    }

    public event EventHandler? CheckedChanged;

    public ToggleSwitch()
    {
        Size = new Size(TW, TH);
        Cursor = Cursors.Hand;
        SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                 ControlStyles.OptimizedDoubleBuffer | ControlStyles.SupportsTransparentBackColor, true);
        BackColor = Color.Transparent;
    }

    protected override void OnMouseEnter(EventArgs e) { _hover = true;  Invalidate(); }
    protected override void OnMouseLeave(EventArgs e) { _hover = false; Invalidate(); }

    protected override void OnClick(EventArgs e)
    {
        _checked = !_checked;
        CheckedChanged?.Invoke(this, EventArgs.Empty);
        Invalidate();
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        var g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;

        Color trackOn  = MainForm.CAccentB;
        Color trackOff = Color.FromArgb(_hover ? 55 : 40, 55, 70);

        // Track
        using var trackPath = MainForm.RoundedPath(0, 0, TW - 1, TH - 1, TH / 2);
        using (var tb = new SolidBrush(_checked ? trackOn : trackOff))
            g.FillPath(tb, trackPath);

        // Thumb
        int thumbX = _checked ? TW - TH + 2 : 2;
        int thumbY = 2;
        int thumbD = TH - 4;
        using var thumbBrush = new SolidBrush(Color.White);
        g.FillEllipse(thumbBrush, thumbX, thumbY, thumbD, thumbD);
    }
}

// ─── ToggleRow (label + ToggleSwitch) ────────────────────────────────────────

internal sealed class ToggleRow : Control
{
    public ToggleSwitch Toggle { get; }
    private readonly string _label;

    public ToggleRow(string label, bool initialValue)
    {
        _label  = label;
        Toggle  = new ToggleSwitch { Checked = initialValue };
        SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                 ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw |
                 ControlStyles.SupportsTransparentBackColor, true);
        BackColor = Color.Transparent;
        Controls.Add(Toggle);
        Toggle.CheckedChanged += (s, e) => Toggle.Checked = Toggle.Checked; // relay
    }

    protected override void OnResize(EventArgs e)
    {
        base.OnResize(e);
        Toggle.Location = new Point(Width - Toggle.Width - 16, (Height - Toggle.Height) / 2);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        var g = e.Graphics;
        g.SmoothingMode     = SmoothingMode.AntiAlias;
        g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
        using var f = new Font("Segoe UI", 9f);
        using var b = new SolidBrush(MainForm.CText);
        g.DrawString(_label, f, b, new PointF(16, (Height - g.MeasureString(_label, f).Height) / 2));
    }
}

// ─── HotkeyBox ────────────────────────────────────────────────────────────────

internal sealed class HotkeyBox : Control
{
    private string  _displayText = Strings.HotkeyNone;
    private bool    _recording;
    private bool    _error;
    private bool    _hoverSet;
    private bool    _hoverClear;

    public event EventHandler? SetClicked;
    public event EventHandler? ClearClicked;

    private const int BtnW  = 58;
    private const int BtnH  = 24;
    private const int BtnGap = 6;

    public HotkeyBox()
    {
        SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                 ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw |
                 ControlStyles.SupportsTransparentBackColor, true);
        BackColor = Color.Transparent;
        Cursor    = Cursors.Default;
    }

    public void StartRecording()
    {
        _recording    = true;
        _error        = false;
        _displayText  = Strings.HotkeyPrompt;
        Invalidate();
    }

    public void SetValue(string text, bool success)
    {
        _recording   = false;
        _error       = !success;
        _displayText = text;
        Invalidate();
    }

    public void CancelRecording(string fallback)
    {
        _recording   = false;
        _error       = false;
        _displayText = fallback;
        Invalidate();
    }

    // ── Hit testing ──────────────────────────────────────────────────────────

    private Rectangle SetRect   => new(Width - (BtnW + BtnGap + BtnW) - 14, (Height - BtnH) / 2, BtnW, BtnH);
    private Rectangle ClearRect => new(Width - BtnW - 14,                   (Height - BtnH) / 2, BtnW, BtnH);

    protected override void OnMouseMove(MouseEventArgs e)
    {
        bool hs = SetRect.Contains(e.Location);
        bool hc = ClearRect.Contains(e.Location);
        if (hs != _hoverSet || hc != _hoverClear) { _hoverSet = hs; _hoverClear = hc; Invalidate(); }
        Cursor = (hs || hc) ? Cursors.Hand : Cursors.Default;
    }

    protected override void OnMouseLeave(EventArgs e)
    { _hoverSet = _hoverClear = false; Invalidate(); }

    protected override void OnMouseUp(MouseEventArgs e)
    {
        if (e.Button != MouseButtons.Left) return;
        if (SetRect.Contains(e.Location))   SetClicked?.Invoke(this, EventArgs.Empty);
        if (ClearRect.Contains(e.Location)) ClearClicked?.Invoke(this, EventArgs.Empty);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        var g = e.Graphics;
        g.SmoothingMode     = SmoothingMode.AntiAlias;
        g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

        // ── Label ─────────────────────────────────────────────────────────────
        using var lf = new Font("Segoe UI", 9f);
        using var lb = new SolidBrush(MainForm.CMuted);
        g.DrawString(Strings.HotkeyLabel, lf, lb, new PointF(16, (Height - g.MeasureString("A", lf).Height) / 2));

        // ── Value display ─────────────────────────────────────────────────────
        int valX = 16 + (int)g.MeasureString(Strings.HotkeyLabel, lf).Width + 8;
        int valW = SetRect.Left - valX - 10;
        var dispColor = _recording ? Color.FromArgb(255, 190, 50)
                      : _error     ? Color.FromArgb(220, 80,  80)
                      : MainForm.CText;
        using var df = new Font("Segoe UI", 9f, _recording ? FontStyle.Italic : FontStyle.Regular);
        using var db = new SolidBrush(dispColor);
        using var dsf = new StringFormat { LineAlignment = StringAlignment.Center, Trimming = StringTrimming.EllipsisCharacter };
        g.DrawString(_displayText, df, db, new RectangleF(valX, 0, valW, Height), dsf);

        // ── Set button ────────────────────────────────────────────────────────
        PaintPillButton(g, SetRect, _recording ? Strings.HotkeyRecording : Strings.HotkeySet,
            _recording ? Color.FromArgb(255, 190, 50) : MainForm.CAccentB, _hoverSet);

        // ── Clear button ──────────────────────────────────────────────────────
        PaintPillButton(g, ClearRect, Strings.HotkeyClear,
            Color.FromArgb(210, 80, 80), _hoverClear);
    }

    private static void PaintPillButton(Graphics g, Rectangle r, string text, Color accent, bool hover)
    {
        var bg = Color.FromArgb(hover ? 35 : 20, accent.R, accent.G, accent.B);
        using var path = MainForm.RoundedPath(r.X, r.Y, r.Width, r.Height, r.Height / 2);
        using (var fb = new SolidBrush(bg)) g.FillPath(fb, path);
        using (var pp = new Pen(Color.FromArgb(hover ? 160 : 100, accent.R, accent.G, accent.B), 1f))
            g.DrawPath(pp, path);
        using var tf  = new Font("Segoe UI", 8f, FontStyle.Bold);
        using var tb  = new SolidBrush(hover ? accent : Color.FromArgb(180, accent.R, accent.G, accent.B));
        using var tsf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
        g.DrawString(text, tf, tb, r, tsf);
    }
}

// ─── Dark menu renderer ───────────────────────────────────────────────────────

internal sealed class DarkMenuRenderer : ToolStripProfessionalRenderer
{
    public DarkMenuRenderer() : base(new DarkColorTable()) { }

    protected override void OnRenderToolStripBackground(ToolStripRenderEventArgs e)
    { using var b = new SolidBrush(MainForm.CCard); e.Graphics.FillRectangle(b, e.AffectedBounds); }

    protected override void OnRenderMenuItemBackground(ToolStripItemRenderEventArgs e)
    { using var b = new SolidBrush(e.Item.Selected ? MainForm.CDim : MainForm.CCard);
      e.Graphics.FillRectangle(b, e.Item.ContentRectangle); }

    protected override void OnRenderItemText(ToolStripItemTextRenderEventArgs e)
    { e.TextColor = e.Item.Enabled ? MainForm.CText : MainForm.CMuted;
      base.OnRenderItemText(e); }

    protected override void OnRenderSeparator(ToolStripSeparatorRenderEventArgs e)
    { int y = e.Item.ContentRectangle.Height / 2;
      using var p = new Pen(MainForm.CBorder);
      e.Graphics.DrawLine(p, e.Item.ContentRectangle.Left, y, e.Item.ContentRectangle.Right, y); }
}

internal sealed class DarkColorTable : ProfessionalColorTable
{
    public override Color MenuBorder                  => MainForm.CBorder;
    public override Color ToolStripDropDownBackground => MainForm.CCard;
    public override Color ImageMarginGradientBegin    => MainForm.CCard;
    public override Color ImageMarginGradientMiddle   => MainForm.CCard;
    public override Color ImageMarginGradientEnd      => MainForm.CCard;
}

// ─── Audio state overlay ──────────────────────────────────────────────────────

internal sealed class OverlayForm : Form
{
    private readonly bool _isMono;
    private readonly System.Windows.Forms.Timer _timer;
    private int _tick;

    private const int FadeInTicks  =  8;
    private const int HoldTicks    = 135;
    private const int FadeOutTicks = 12;
    private const int TotalTicks   = FadeInTicks + HoldTicks + FadeOutTicks;
    private const int W = 240, H = 78;

    public OverlayForm(bool isMono)
    {
        _isMono         = isMono;
        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar   = false;
        TopMost         = true;
        StartPosition   = FormStartPosition.Manual;
        BackColor       = Color.Black;
        TransparencyKey = Color.Black;
        Size            = new Size(W, H);
        Opacity         = 0;

        var wa = Screen.FromPoint(Cursor.Position).WorkingArea;
        Location = new Point(wa.Right - W - 24, wa.Bottom - H - 24);

        SetStyle(ControlStyles.OptimizedDoubleBuffer |
                 ControlStyles.AllPaintingInWmPaint  |
                 ControlStyles.UserPaint, true);

        _timer = new System.Windows.Forms.Timer { Interval = 20 };
        _timer.Tick += (_, _) =>
        {
            _tick++;
            Opacity = _tick <= FadeInTicks                 ? (double)_tick / FadeInTicks
                    : _tick <= FadeInTicks + HoldTicks     ? 1.0
                    : _tick <= TotalTicks                  ? 1.0 - (double)(_tick - FadeInTicks - HoldTicks) / FadeOutTicks
                    : 0;
            if (_tick > TotalTicks) { _timer.Stop(); Close(); }
        };
        _timer.Start();
    }

    protected override bool ShowWithoutActivation => true;
    protected override CreateParams CreateParams
    {
        get { var cp = base.CreateParams; cp.ExStyle |= 0x08000000 | 0x00000080; return cp; }
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        var g = e.Graphics;
        g.SmoothingMode     = SmoothingMode.AntiAlias;
        g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

        Color accent = _isMono ? MainForm.CAccentB : MainForm.CAccentG;

        using var path = MainForm.RoundedPath(0, 0, W - 1, H - 1, 12);
        using (var bg = new SolidBrush(Color.FromArgb(20, 20, 28))) g.FillPath(bg, path);
        using (var bp = new Pen(accent, 1.2f)) g.DrawPath(bp, path);

        // Circle badge
        const int cx = 16, cs = 46;
        int cty = (H - cs) / 2;
        var badgeBg = Color.FromArgb(30, accent.R, accent.G, accent.B);
        using (var bb = new SolidBrush(badgeBg)) g.FillEllipse(bb, cx, cty, cs, cs);
        using (var bp2 = new Pen(accent, 1.5f)) g.DrawEllipse(bp2, cx, cty, cs, cs);

        using var lf  = new Font("Segoe UI", 22f, FontStyle.Bold, GraphicsUnit.Pixel);
        using var lb  = new SolidBrush(accent);
        using var lsf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
        g.DrawString(_isMono ? "M" : "S", lf, lb, new RectangleF(cx, cty, cs, cs), lsf);

        int tx = cx + cs + 14, mid = H / 2;
        using var mf = new Font("Segoe UI", 20f, FontStyle.Bold, GraphicsUnit.Pixel);
        using var mb = new SolidBrush(accent);
        g.DrawString(_isMono ? Strings.BtnMono : Strings.BtnStereo, mf, mb, new PointF(tx, mid - 16));

        using var sf = new Font("Segoe UI", 10f, GraphicsUnit.Pixel);
        using var sb = new SolidBrush(MainForm.CMuted);
        g.DrawString(_isMono ? Strings.OverlayMonoSub : Strings.OverlayStereoSub, sf, sb, new PointF(tx + 1, mid + 5));
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing) _timer.Dispose();
        base.Dispose(disposing);
    }
}

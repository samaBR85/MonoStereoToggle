using System.Globalization;

namespace MonoStereoToggle;

/// <summary>
/// All UI strings — returns PT or EN based on the system UI culture.
/// Add new languages by extending the ternary chains.
/// </summary>
internal static class Strings
{
    private static bool IsPt =>
        CultureInfo.CurrentUICulture.TwoLetterISOLanguageName == "pt";

    // ── Main window ───────────────────────────────────────────────────────────
    public static string AppTitle      => IsPt ? "Alternador de Áudio"    : "Audio Switcher";
    public static string StatusMono    => IsPt ? "Mono ativo  —  canais L+R mixados em um" : "Mono active  —  L+R channels mixed";
    public static string StatusStereo  => IsPt ? "Estéreo ativo  —  canais independentes"  : "Stereo active  —  independent channels";
    public static string StatusWait    => IsPt ? "Aguarde…"               : "Please wait…";
    public static string BtnStereo     => IsPt ? "ESTÉREO"                : "STEREO";
    public static string BtnMono       => "MONO";

    // ── Device list ───────────────────────────────────────────────────────────
    public static string DevicesLabel  => IsPt ? "Dispositivos de saída disponíveis:"       : "Available output devices:";
    public static string DeviceHint    => IsPt ? "Áudio mono é uma configuração global do Windows" : "Mono audio is a global Windows setting";
    public static string DeviceDefault => IsPt ? "(padrão)"               : "(default)";

    // ── Checkboxes / footer ───────────────────────────────────────────────────
    public static string ChkStartup    => IsPt ? "  Iniciar com o Windows"         : "  Start with Windows";
    public static string ChkTray       => IsPt ? "  Iniciar na bandeja do sistema" : "  Start minimized to tray";
    public static string Footer        => IsPt ? "Clique duplo no ícone da bandeja para abrir  |  Atalho funciona globalmente"
                                               : "Double-click tray icon to open  |  Shortcut works globally";

    // ── Hotkey ────────────────────────────────────────────────────────────────
    public static string HotkeyLabel     => IsPt ? "Atalho:"                          : "Shortcut:";
    public static string HotkeyNone      => IsPt ? "Nenhum"                           : "None";
    public static string HotkeyRecording => IsPt ? "Pressione…"                       : "Press…";
    public static string HotkeyPrompt    => IsPt ? "Pressione a combinação de teclas…" : "Press a key combination…";
    public static string HotkeyInUse     => IsPt ? "Combinação já em uso"             : "Combination already in use";
    public static string HotkeySet       => IsPt ? "Definir"                          : "Set";
    public static string HotkeyClear     => IsPt ? "Limpar"                           : "Clear";
    public static string KeySpace        => IsPt ? "Espaço"                           : "Space";

    // ── Tray ──────────────────────────────────────────────────────────────────
    public static string TrayMono      => IsPt ? "Áudio MONO"        : "MONO Audio";
    public static string TrayStereo    => IsPt ? "Áudio ESTÉREO"     : "STEREO Audio";
    public static string TrayToMono    => IsPt ? "Mudar para Mono"   : "Switch to Mono";
    public static string TrayToStereo  => IsPt ? "Mudar para Estéreo": "Switch to Stereo";
    public static string TrayOpen      => IsPt ? "Abrir"             : "Open";
    public static string TrayClose     => IsPt ? "Fechar"            : "Close";

    // ── Errors ────────────────────────────────────────────────────────────────
    public static string ErrorPrefix   => IsPt ? "Erro:"                       : "Error:";
    public static string StartupError  => IsPt ? "Erro ao configurar startup:" : "Startup configuration error:";

    // ── Overlay ───────────────────────────────────────────────────────────────
    public static string OverlayMonoSub   => IsPt ? "Áudio mono ativo"    : "Mono audio active";
    public static string OverlayStereoSub => IsPt ? "Áudio estéreo ativo" : "Stereo audio active";
}

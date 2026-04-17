using System.Threading;
using System.Windows.Forms;

namespace MonoStereoToggle;

internal static class Program
{
    [STAThread]
    static void Main(string[] args)
    {

        using var mutex = new Mutex(true, "Global\\MonoStereoToggle_SingleInstance", out bool created);
        if (!created) return;

        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.SetHighDpiMode(HighDpiMode.PerMonitorV2);

        Application.Run(new MainForm(args.Contains("--tray")));
    }
}

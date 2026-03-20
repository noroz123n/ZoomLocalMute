using System;
using System.Windows.Forms;

namespace ZoomLocalMute
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            using var trayApp = new TrayApplication();
            Application.Run();
        }
    }
}
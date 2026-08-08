using System;
using System.Diagnostics;
using System.Security.Principal;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SteamRouteTool
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        private static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            Application.ThreadException += (sender, e) => ReportFatal(e.Exception);
            AppDomain.CurrentDomain.UnhandledException += (sender, e) => ReportFatal(e.ExceptionObject as Exception);
            TaskScheduler.UnobservedTaskException += (sender, e) =>
            {
                Debug.WriteLine("Unobserved task exception: " + e.Exception);
                e.SetObserved();
            };

            WarnIfNotElevated();

            int appId = PromptForAppId();
            if (appId <= 0) return; // The user cancelled the prompt.

            Application.Run(new MainForm(appId));
        }

        /// <summary>Asks which game to work on. Returns 0 when the user cancels.</summary>
        private static int PromptForAppId()
        {
            using (var prompt = new AppIdPromptForm(AppSettings.InitialAppId))
            {
                return prompt.ShowDialog() == DialogResult.OK ? prompt.AppId : 0;
            }
        }

        /// <summary>
        /// The manifest already asks for elevation; this only catches the case where the
        /// manifest has been stripped, which would otherwise fail later with an opaque
        /// access denied from the firewall.
        /// </summary>
        private static void WarnIfNotElevated()
        {
            if (IsElevated()) return;

            MessageBox.Show(
                "SteamRouteTool is not running as administrator, so it cannot add or remove firewall rules." +
                Environment.NewLine + Environment.NewLine +
                "Close it and start it again with \"Run as administrator\".",
                "Administrator rights required",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }

        private static bool IsElevated()
        {
            try
            {
                using (WindowsIdentity identity = WindowsIdentity.GetCurrent())
                {
                    return new WindowsPrincipal(identity).IsInRole(WindowsBuiltInRole.Administrator);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Could not determine elevation: " + ex);
                return true; // Do not nag when the check itself is unavailable.
            }
        }

        private static void ReportFatal(Exception error)
        {
            if (error == null) return;

            Debug.WriteLine("Unhandled exception: " + error);
            MessageBox.Show(
                error.Message,
                "SteamRouteTool encountered a problem",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }
}

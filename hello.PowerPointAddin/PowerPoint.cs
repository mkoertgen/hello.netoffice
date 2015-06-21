using System;
using System.Diagnostics;
using System.Reflection;
using System.Security.Principal;
using System.Windows.Forms;
using NetOffice.OfficeApi.Enums;
using NetOffice.PowerPointApi.Tools;
using Application = NetOffice.PowerPointApi.Application;

namespace hello.PowerPointAddin
{
    static class PowerPoint
    {
        public static void StartRegistered<TAddin>() where TAddin : COMAddin
        {
            var pptApp = Application.GetActiveInstance();

            if (!Com.IsRegistered<TAddin>())
            {
                var isAdmin = IsAdministrator();
                try
                {
                    Com.Register<MyAddin>();
                    if (isAdmin) return;
                }
                catch (UnauthorizedAccessException)
                {
                    if (isAdmin) throw;
                    RestartAsAdmin();
                }

                if (pptApp != null)
                {
                    if (MessageBox.Show(null,
                        $"PowerPoint must be restarted in order for the AddIn '{typeof(TAddin).Name}' to load. Restart PowerPoint now?",
                        "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    { 
                        pptApp.Quit();
                        pptApp.Dispose();
                        pptApp = null;
                    }
                }
            }

            // start PowerPoint (again)
            if (pptApp == null)
                pptApp = new Application();
            // bring to front
            pptApp.Visible = MsoTriState.msoTrue;
        }

        private static bool IsAdministrator()
        {
            return new WindowsPrincipal(WindowsIdentity.GetCurrent())
                .IsInRole(WindowsBuiltInRole.Administrator);
        }

        private static void RestartAsAdmin()
        {
            // Restart program and run as admin
            var exeName = new Uri(Assembly.GetEntryAssembly().CodeBase).LocalPath;
            var startInfo = new ProcessStartInfo(exeName) { Verb = "runas" };
            var process = Process.Start(startInfo);
            process.EnableRaisingEvents = true;
            process.WaitForExit();
        }
    }
}
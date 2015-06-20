using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using NetOffice.OfficeApi;
using NetOffice.PowerPointApi.Tools;
using NetOffice.Tools;
using NLog;
using NLog.Config;
using NLog.Targets;

namespace hello.PowerPointAddin
{
    [COMAddin("MyAddin", "A sample power point addin", 3)]
    [Guid("4C454497-F5C1-4429-A5B0-1D3D99199A52"), ProgId("hello.PowerPointAddin.MyAddin"), Tweak(true)]
    // loadbehavior: cf.: https://msdn.microsoft.com/en-us/library/bb386106.aspx#LoadBehavior
    // command line safe, cf.: https://msdn.microsoft.com/en-us/library/19dax6cz.aspx
    [CustomUI("hello.PowerPointAddin.RibbonUI.xml"), RegistryLocation(RegistrySaveLocation.CurrentUser)]
    [CustomPane(typeof(SamplePane), "Sample Pane", true, PaneDockPosition.msoCTPDockPositionTop, PaneDockPositionRestrict.msoCTPDockPositionRestrictNoChange, 50, 50)]
    public class MyAddin : COMAddin
    {
        public MyAddin()
        {
            InitializeLogging();

            // Enable shared debug output and send a load message(use NOTools.ConsoleMonitor.exe to observe the shared console output)
            Factory.Console.EnableSharedOutput = true;
            Factory.Console.SendPipeConsoleMessage(null, "Addin has been loaded.");

            // We want observe the current count of open proxies with NOTools.ConsoleMonitor.exe 
            Factory.Settings.EnableProxyCountChannel = true;

            // trigger the well known IExtensibility2 methods, this is very similar to VSTO
            OnStartupComplete += Addin_OnStartupComplete;
        }

        // ouer ribbon instance to manipulate ui at runtime 
        private IRibbonUI RibbonUi { get; set; }

        // attached in ctor to say hello in console
        private void Addin_OnStartupComplete(ref Array custom)
        {
            // You see the host application is accessible as property from the class instance.
            // The application property was disposed automaticly while shutdown.
            Factory.Console.WriteLine("Host Application Version is:{0}.", Application.Version);
        }

        // defined in RibbonUI.xml to get a instance for ouer ribbon ui.
        public void OnLoadRibbonUi(IRibbonUI ribbonUi)
        {
            RibbonUi = ribbonUi;
        }

        // defined in RibbonUI.xml to make sure the checkbutton state is up-to-date and synchronized with taskpane visibility.
        public bool OnGetPressedPanelToggle(IRibbonControl control)
        {
            return TaskPanes[0].Visible;
        }

        // defined in RibbonUI.xml to track the user clicked ouer checkbutton. then we upate the panel visibility at hand.
        public void OnCheckPanelToggle(IRibbonControl control, bool pressed)
        {
            TaskPanes[0].Visible = pressed;
        }

        // defined in RibbonUI.xml to catch the user click for the about button
        public void OnClickAboutButton(IRibbonControl control)
        {
            MessageBox.Show("Sample Addin", Type.FullName);
        }

        /*
        * Now you see the way to exend or modify the register/unregister process if you want.
        * We define 2 static methods with the RegisterFunction attribute, we use CallBeforeAndAfter as condition.
        * This condition means the register method in the base class call our method as first (before registry modification) and as last(after registry modification).
        * The register call argument give you the info what is it currently. Replace means the method in the base class does nothing and its your task to create the registry keys.
        * Same thing with Unregister method. 
        */

        [RegisterFunction(RegisterMode.CallBeforeAndAfter)]
        public static void Register(Type type, RegisterCall registerCall)
        {
            switch (registerCall)
            {
                case RegisterCall.CallAfter:
                    break;
                case RegisterCall.CallBefore:
                    break;
                case RegisterCall.Replace:
                    break;
                default:
                    break;
            }
        }

        [UnRegisterFunction(RegisterMode.CallBeforeAndAfter)]
        public static void UnRegister(Type type, RegisterCall registerCall)
        {
            switch (registerCall)
            {
                case RegisterCall.CallAfter:
                    break;
                case RegisterCall.CallBefore:
                    break;
                case RegisterCall.Replace:
                    break;
                default:
                    break;
            }
        }


        /*
         * At last you see some options for troubleshooting. The COMAddin base class is not a blackbox.
        */

        // This error handler is used for IExtensibility2 events (your code) and the COMAddin methods GetCustomUI, CTPFactoryAvailable and CreateFactory(also overwrites).
        // the first argument shows in which method the error is occured. The second argument is the detailed exception info. 
        // Rethrow the exception otherwise the exception is marked as handled.
        protected override void OnError(ErrorMethodKind methodKind, Exception exception)
        {
            TraceAndPresent($"An error occured in '{methodKind}'.", exception);
        }

        // This method demonstrate an error handler for the register/unregister process.
        // For example you have an security issues while register or something like that then you can implement a static errorhandler method.
        // The first argument shows you the error occurs in Register or Unregister.
        // The second argument is the thrown exception. Rethrow the exception to signalize an error to the environment otherwise the exception is handled.
        [RegisterErrorHandler]
        public static void RegisterErrorHandler(RegisterErrorMethodKind methodKind, Exception exception)
        {
            TraceAndPresent($"An error occured in '{methodKind}'.", exception);
        }

        private static void TraceAndPresent(string msg, Exception exception)
        {
            Trace.TraceError("{0}: {1}", msg, exception);
            MessageBox.Show(msg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private static void InitializeLogging(string logConfigFileName = null)
        {
            Trace.Listeners.Add(new NLogTraceListener());

            var logconfig = logConfigFileName ?? "NLog.config";
            var path = Environment.ExpandEnvironmentVariables(logconfig);
            var logConfigFile = new FileInfo(path);
            if (logConfigFile.Exists)
                LogManager.Configuration = new XmlLoggingConfiguration(logConfigFile.FullName);
            else
            {
                LogManager.Configuration = DefaultLogging(ModeDetector.IsDebug ? LogLevel.Debug : LogLevel.Info);
                // NOTE that we log the warning AFTER initializing logging!
                Trace.TraceWarning("Unable to find log configuration file: '{0}'. Using default configuration", logconfig);
            }
        }

        private static LoggingConfiguration DefaultLogging(LogLevel logLevel = null)
        {
            // cf.: http://stackoverflow.com/questions/24070349/nlog-switching-from-nlog-config-to-programmatic-configuration
            var config = new LoggingConfiguration();

            var fileTarget = new FileTarget
            {
                Layout = "${longdate} [${threadid}] ${uppercase:${level}} - ${message}",
                FileName = string.Format("${{environment:LOCALAPPDATA}}/{0}/{0}/Logs/{1}_${{environment:USERNAME}}.${{environment:USERDOMAIN}}.log",
                    nameof(MyAddin), nameof(MyAddin)),
                Header = "[Open Log]",
                Footer = "[Close Log]",
                ArchiveFileName = string.Format("${{environment:LOCALAPPDATA}}/{0}/{0}/Logs/{1}_${{environment:USERNAME}}.${{environment:USERDOMAIN}}.{{#}}.log",
                    nameof(MyAddin), nameof(MyAddin)),
                ArchiveAboveSize = 1048576,
                ArchiveEvery = FileArchivePeriod.None,
                ArchiveNumbering = ArchiveNumberingMode.Rolling,
                MaxArchiveFiles = 5,
                ConcurrentWrites = false,
                KeepFileOpen = true,
                Encoding = Encoding.UTF8
            };

            config.AddTarget("f", fileTarget);

            var rule = new LoggingRule("*", logLevel ?? LogLevel.Debug, fileTarget);
            config.LoggingRules.Add(rule);

            return config;
        }

    }
}

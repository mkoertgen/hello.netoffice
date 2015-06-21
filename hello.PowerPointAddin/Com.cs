using System;
using System.Reflection;
using System.Runtime.InteropServices;
using NetOffice.PowerPointApi.Tools;

namespace hello.PowerPointAddin
{
    static class Com
    {
        static readonly RegistrationServices Registrar = new RegistrationServices();

        public static bool IsRegistered<TAddin>() where TAddin : COMAddin
        {
            var progId = typeof(TAddin).GetCustomAttribute<ProgIdAttribute>().Value;
            return Type.GetTypeFromProgID(progId) == typeof(TAddin);
            // TODO: May need to update registration for newer version
        }

        public static bool Register<TAddin>() where TAddin : COMAddin
        {
            return Registrar.RegisterAssembly(typeof(TAddin).Assembly, AssemblyRegistrationFlags.SetCodeBase);
        }

        public static bool Unregister<TAddin>() where TAddin : COMAddin
        {
            return Registrar.UnregisterAssembly(typeof(TAddin).Assembly);
        }

    }
}
using System.Diagnostics.CodeAnalysis;

// ReSharper disable once CheckNamespace
namespace System.Diagnostics
{
    /// <summary>
    /// Used to detect the current build compiling symbols
    /// </summary>
    [ExcludeFromCodeCoverage]
    static class ModeDetector
    {
        public static bool IsDebug
        {
            get
            {
#if (DEBUG)
                return true;
#else
                return false;
#endif
            }
        }

        public static bool IsTrace
        {
            get
            {
#if (TRACE)
                return true;
#else
                return false;
#endif
            }
        }

        public static bool IsRelease => !IsDebug;
    }
}
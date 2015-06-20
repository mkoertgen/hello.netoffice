using System.Reflection;
using FluentAssertions;
using NetOffice.Tools;
using NUnit.Framework;

namespace hello.PowerPointAddin
{
    [TestFixture]
    class MyAddinTest
    {
        [Test]
        public void Test_CustomUi_resource()
        {
            var customUi = typeof (MyAddin).GetCustomAttribute<CustomUIAttribute>().Value;
            using (var stream = typeof (MyAddin).Assembly.GetManifestResourceStream(customUi))
            {
                stream.Should().NotBeNull();
            }
        }
    }
}
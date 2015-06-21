namespace hello.PowerPointAddin
{
    static class Program
    {
        public static void Main(string[] args)
        {
            PowerPoint.StartRegistered<MyAddin>();
        }
    }
}
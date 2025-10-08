using System;
using System.Reflection;

namespace NET8AutomatedReports
{
    internal class CheckAssemblyInfo
    {
        public static void ShowCopyright()
        {
            var assembly = Assembly.GetExecutingAssembly();
            var copyright =
                assembly.GetCustomAttribute<AssemblyCopyrightAttribute>();

            Console.WriteLine("Copyright: " + (copyright?.Copyright ?? "No encontrado"));
        }
    }
}

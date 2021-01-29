using Autofac;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;
using WinDev.Environment;

namespace WinDev.OpenXmlDemonstrator
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            ApplicationStarter.CreateHostContainer(Singletons);
            Application.Run(new DirectoryHierarchyExporter());
        }

        /// <summary>
        /// Place to put application specific dependency registration 
        /// </summary>
        /// <param name="builder"></param>
        static void Singletons(ContainerBuilder builder)
        {
            string baseDirectoryPath = AppDomain.CurrentDomain.BaseDirectory + "bin";
            if (!Directory.Exists(baseDirectoryPath))
                baseDirectoryPath = AppDomain.CurrentDomain.BaseDirectory;

            var assemblies = Directory.EnumerateFiles(baseDirectoryPath, "*.dll", SearchOption.TopDirectoryOnly)
                .Where(filePath => Path.GetFileName(filePath).StartsWith("WinDev"))
                .Select(Assembly.LoadFrom).Where(assemblyType =>
                assemblyType.FullName.StartsWith("WinDev") && !assemblyType.FullName.Contains("") &&
                !assemblyType.FullName.Contains("WinDev.OpenXmlDemonstrator")
                ).ToArray();

            builder.RegisterAssemblyTypes(assemblies)
            .AsImplementedInterfaces().InstancePerLifetimeScope();
        }
    }
}

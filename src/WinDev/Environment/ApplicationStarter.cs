using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autofac;
using Autofac.Configuration;
using WinDev.Environment.AutofacUtil;
using WinDev.Logging;

namespace WinDev.Environment
{
    public static class ApplicationStarter
    {
        public static IContainer CreateHostContainer(Action<ContainerBuilder> registrations)
        {
            var builder = new ContainerBuilder();
            // Application paths and parameters
            builder.RegisterModule(new LoggingModule());



            registrations(builder);

            var autofacSection = ConfigurationManager.GetSection(ConfigurationSettingsReaderConstants.DefaultSectionName);
            if (autofacSection != null)
                builder.RegisterModule(new ConfigurationSettingsReader());

            var optionalHostConfig = @"Config\Host.config";
            if (File.Exists(optionalHostConfig))
                builder.RegisterModule(new ConfigurationSettingsReader(ConfigurationSettingsReaderConstants.DefaultSectionName, optionalHostConfig));

            //var optionalComponentsConfig = @"Config\HostComponents.config";
            //if (File.Exists(optionalComponentsConfig))
            //    builder.RegisterModule(new HostComponentsConfigModule(optionalComponentsConfig));

            var container = builder.Build();        

            return container;
        }

    }
}

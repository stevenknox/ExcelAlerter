using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Hosting.Internal;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.FileProviders;
using Serilog;

namespace ExcelAlerter
{
    public static class AppBootstrapper
    {
        public static ServiceProvider Bootstrap(string basePath)
        {
            var configuration = new ConfigurationBuilder()
                    .SetBasePath(basePath)
                    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                    .Build();

            var appSettings = new AppSettings();
            configuration.GetSection("App").Bind(appSettings);

            Log.Logger = new LoggerConfiguration()
                .ReadFrom.Configuration(configuration)
                .CreateLogger();

            var serviceCollection = new ServiceCollection();

            ConfigureServices(serviceCollection, configuration, basePath);

            var serviceProvider = serviceCollection.BuildServiceProvider();
            return serviceProvider;
        }

        private static void ConfigureServices(IServiceCollection services, IConfiguration config, string basePath)
        {
            services.AddLogging(configure => configure.AddSerilog());

            services.AddSingleton(config);
            services.AddScoped<IAlerter, Alerter>();
            services.Configure<AppSettings>(config.GetSection("App"));

            services.AddSingleton<IHostingEnvironment, HostingEnvironment>(sp =>
             {
                return new HostingEnvironment {
                     ContentRootPath = basePath,
                    ContentRootFileProvider = new PhysicalFileProvider(basePath)
                };
             });

             System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        }
    }
}
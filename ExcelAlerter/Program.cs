using static System.Console;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Serilog;
using System.IO;

namespace ExcelAlerter
{
    class Program
    {
        static void Main(string[] args)
        {
            ServiceProvider serviceProvider = Bootstrap();

            var alerter = serviceProvider.GetService<IAlerter>();

            alerter.Notify();

            ReadLine();

        }

        private static ServiceProvider Bootstrap()
        {
            var configuration = new ConfigurationBuilder()
                    .SetBasePath(Directory.GetCurrentDirectory())
                    .AddJsonFile("appsettings.json")
                    .Build();

            Log.Logger = new LoggerConfiguration()
                .ReadFrom.Configuration(configuration)
                .CreateLogger();

            var serviceCollection = new ServiceCollection();

            ConfigureServices(serviceCollection, configuration);

            var serviceProvider = serviceCollection.BuildServiceProvider();
            return serviceProvider;
        }

        private static void ConfigureServices(IServiceCollection services, IConfiguration configuration)
        {
            services.AddLogging(configure => configure.AddSerilog());

            services.AddScoped<IAlerter, Alerter>();
        }
    }
}

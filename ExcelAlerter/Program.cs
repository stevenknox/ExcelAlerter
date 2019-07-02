using static System.Console;
using Microsoft.Extensions.DependencyInjection;
using System.IO;

namespace ExcelAlerter
{
    class Program
    {
        static void Main(string[] args)
        {
            ServiceProvider serviceProvider = AppBootstrapper.Bootstrap(Directory.GetCurrentDirectory());

            var alerter = serviceProvider.GetService<IAlerter>();

            alerter.LoadData();

            // WriteLine(alerter.Data.ToStringTable());
            WriteLine("Press any key to exit.");
            ReadLine();
        }


    }
}

using Microsoft.Extensions.DependencyInjection;
using System.IO;
using ExcelAlerter;
using Xunit;
using Microsoft.Extensions.Options;
using Shouldly;


namespace Tests
{
    public class AlerterTests
    {
        private readonly IAlerter alerter;
        private readonly AppSettings settings; 
        public AlerterTests()
        {
           var serviceProvider = AppBootstrapper.Bootstrap(Directory.GetCurrentDirectory() + "../../../../../ExcelAlerter");

           alerter = serviceProvider.GetService<IAlerter>();
           settings= serviceProvider.GetService<IOptionsMonitor<AppSettings>>().CurrentValue;
        }


        [Fact]
        public void Should_Have_Valid_Config()
        {
            settings.FilePath.ShouldContain("Calibration - First Responders.xlsx");
            settings.ExcelDateField.ShouldBe("Expiry Date");
            settings.DaysInAdvance.ShouldBe(7);
        }

        [Fact]
        public void Should_Load_Data_From_Excel()
        {
            alerter.LoadData();

            alerter.Data.ShouldNotBeEmpty();
        }
    }
}

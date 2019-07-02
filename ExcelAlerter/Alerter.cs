using Microsoft.Extensions.Logging;

namespace ExcelAlerter
{
    public interface IAlerter
    {
        void Notify();
    }
    public class Alerter : IAlerter
    {
        private readonly ILogger<Alerter> _log;

        public Alerter(ILogger<Alerter> log)
        {
            _log = log;
        }

        public void Notify()
        {
            _log.LogInformation("Notifying");
        }
    }
}
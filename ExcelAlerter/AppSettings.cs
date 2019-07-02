namespace ExcelAlerter
{
    public class AppSettings
    {
        public string FilePath { get; set; }
        public string ExcelDateField { get; set; }
        public int DaysInAdvance { get; set; }
        public string[] Worksheets { get; set; }
    }
}
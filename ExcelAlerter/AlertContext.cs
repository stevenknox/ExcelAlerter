using System;
using System.Collections.Generic;

namespace ExcelAlerter
{
    public class AlertContext
    {
        public string Id { get; set; }
        public string Who { get; set; }
        public string What { get; set; }
        public string When { get; set; }
        public DateTime AlertDate { get; set; }
    }
}
using System;
using System.Collections.Generic;

namespace ExcelAlerter
{

    public class Equipment
    {
        public static Equipment Create(string id, string name, DateTime expiry) => new Equipment
        {
            Id = id,
            Name = name,
            Expiry = expiry
        };

        public string Id { get; set; }
        public string Name { get; set; }
        public DateTime Expiry { get; set; }
    }
}
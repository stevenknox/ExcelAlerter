using System;
using System.Collections.Generic;

namespace ExcelAlerter
{
    public class Engineer
    {
        public string Name { get; set; }
        public string Role { get; set; }
        public List<Equipment> Equipment { get; set; } = new List<Equipment>();
        public static Engineer Create(string name, string role) => new Engineer
        {
            Name = name,
            Role = role
        };

        public Engineer AddEquipment(Equipment equipment)
        {
            Equipment.Add(equipment);
            return this;
        }
    }
}
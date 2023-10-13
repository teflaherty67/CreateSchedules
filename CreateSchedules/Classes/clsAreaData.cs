using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateSchedules
{
    internal class clsAreaData
    {
        public string Number { get; set; }
        public string Name { get; set; }

        public clsAreaData(string number, string name)
        {
            Number = number;
            Name = name;
        }
    }
}

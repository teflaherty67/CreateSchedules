using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateSchedules
{
    internal class clsFloorAreaData
    {
        public string Number { get; set; }
        public string Name { get; set; }
        public string Category { get; set; }
        public string Comments { get; set; }

        public clsFloorAreaData(string number, string name, string category, string comments)
        {
            Number = number;
            Name = name;
            Category = category;
            Comments = comments;
        }
    }
}

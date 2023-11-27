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
        public string Category { get; set; }
        public string Comments { get; set; }
        public int Ratio { get; set; }

        public clsAreaData(string number, string name, string category, string comments)
        {
            Number = number;
            Name = name;
            Category = category;
            Comments = comments;
            Ratio = 99;
        }
        public clsAreaData(string number, string name, int ratio)
        {
            Number = number;
            Name = name;
            Ratio = ratio;
        }
    }
}

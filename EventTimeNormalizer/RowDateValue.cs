using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EventTimeNormalizer
{
    public class RowDateValue
    {
        public RowDateValue(int OriginalPosition)
        {
            this.OriginalPosition = OriginalPosition;
        }
        public int OriginalPosition { get; set; }
        public double Value { get; set; }
        public DateTime Date { get; set; }

        public string[] Items { get; set; } 
    }
}

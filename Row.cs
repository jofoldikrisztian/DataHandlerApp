using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MartinAppGUI
{
    class Row
    {
        public int id { get; set; }
        public string targyKod { get; set; }
        public string targyNev { get; set; }
        public int letszam { get; set; }
        public List<string> oktatok { get; set; }

        public Row()
        {
            oktatok = new List<string>();
        }

    }
}

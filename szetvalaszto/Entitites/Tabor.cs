using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace szetvalaszto
{
    public class Tabor
    {
        public List<Par> parok;
        public int letszam;
        public Tabor(int letszam, List<Par> parok)
        {
            this.letszam = letszam;
            this.parok = parok;
        }

        public Tabor(int letszam)
        {
            this.letszam = letszam;
            this.parok = new List<Par>();
        }
    }
}

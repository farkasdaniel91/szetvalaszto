using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace szetvalaszto
{
    public class Par
    {
        public int kaszt;
        public string par;
        public List<Preferencia> preferenciak;

        public Par(int evfolyam, string par)
        {
            this.kaszt = evfolyam;
            this.par = par;
        }
        public Par(string par)
        {
            this.par = par;
        }

        public Par(List<Preferencia> pref, int evfolyam, string par)
        {
            this.preferenciak = pref;
            this.kaszt = evfolyam;
            this.par = par;
        }
    }
}

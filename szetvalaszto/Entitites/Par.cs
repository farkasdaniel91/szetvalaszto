using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace szetvalaszto
{
    public class Par
    {
        public int evfolyam;
        public string par;
        public Par(int evfolyam, string par)
        {
            this.evfolyam = evfolyam;
            this.par = par;
        }
        public Par(string par)
        {
            this.par = par;
        }
    }
}

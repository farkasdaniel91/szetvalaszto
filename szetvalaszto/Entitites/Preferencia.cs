using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace szetvalaszto
{
    public class Preferencia
    {
        public string valaszto;
        public string valasztott;
        public int prefpont;
        public string key;
        public int evfolyam;

        public Preferencia(string valaszto, string valasztott, int prefpont, string key)
        {
            this.valaszto = valaszto;
            this.valasztott = valasztott;
            this.prefpont = prefpont;
            this.key = key;
        }
        public Preferencia(string valaszto, string valasztott, int prefpont, int evfolyam)
        {
            this.valaszto = valaszto;
            this.valasztott = valasztott;
            this.prefpont = prefpont;
            this.evfolyam = evfolyam;
        }
    }
}

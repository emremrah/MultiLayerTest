using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nTierApplicationExample.ValuesLayer
{
    public class Sehir : Entity
    {
        public int id { get; set; }
        public string ad { get; set; }

        public Sehir (int id, string ad)
        {
            this.id = id;
            this.ad = ad;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nTierApplicationExample.ValuesLayer
{
    public class Ilce : Entity
    {
        public int id { get; set; }
        public string ad { get; set; }
        public int sehirId { get; set; }

        public Ilce (int id, string ad, int sehirId)
        {
            this.id = id;
            this.ad = ad;
            this.sehirId = sehirId;
        }
    }
}

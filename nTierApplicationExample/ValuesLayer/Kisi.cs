using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nTierApplicationExample.ValuesLayer
{
    public class Kisi : Entity
    {
        public int id { get; set; }
        public string ad { get; set; }
        public string soyad { get; set; }
        public int yas { get; set; }
        public string adres { get; set; }
        public string sehir { get; set; }
        public string ilce { get; set; }

        public Kisi (int id, string ad, string soyad, int yas, string adres, string sehir, string ilce)
        {
            this.id = id;
            this.ad = ad;
            this.soyad = soyad;
            this.yas = yas;
            this.sehir = sehir;
            this.adres = adres;
            this.ilce = ilce;
        }
    }
}

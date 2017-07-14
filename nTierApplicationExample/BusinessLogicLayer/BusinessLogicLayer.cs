using nTierApplicationExample.ValuesLayer;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DAL = nTierApplicationExample.DataAccessLayer.DataAccessLayer;


namespace nTierApplicationExample.BusinessLogicLayer
{
    class BusinessLogicLayer
    {
        //DAL = DATA ACCESS LAYER
        //SELECT
        public DataTable Get (string type)
        {
            try {
                DAL DAL = new DAL();
                return DAL.Read(type);
            } catch {
                throw;
            }
        }

        //INSERT
        public void Insert<T> (T newEntity) where T : Entity
        {
            try {
                DAL DAL = new DAL();
                DAL.Insert<T>(newEntity);
            } catch {
                throw;
            }
        }

        //UPDATE
        public void Update<T> (T newEntity) where T : Entity
        {
            try {
                DAL DAL = new DAL();
                DAL.Update<T>(newEntity);
            } catch {
                throw;
            }
        }

        //DELETE
        public void DeleteSehir (Sehir sehir)
        {
            try {
                DAL DAL = new DAL();
                DAL.Delete(sehir);
            } catch {
                throw;
            }
        }
    }
}

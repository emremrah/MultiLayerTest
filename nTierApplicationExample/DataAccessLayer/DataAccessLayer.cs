using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Data;
using nTierApplicationExample.ValuesLayer;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop;

namespace nTierApplicationExample.DataAccessLayer
{
    class DataAccessLayer
    {
        public string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\calisma\kisi.xlsx; Extended Properties='Excel 12.0 xml;HDR=YES;'";
        
        OleDbConnection connection = new OleDbConnection();
        DataTable dataTable = new DataTable();  //Bellekteki verilerin bir tablosunu temsil ediyor.

        //SELECT
        public DataTable Read(string type)
        {
            connection.ConnectionString = connectionString;
            if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                connection.Open();

            OleDbCommand command = new OleDbCommand("SELECT * FROM ["+type+"$]", connection);
            return Execute(command);
            //try
            //{
            //    if (type == "Sehir")
            //    {
            //        /* Veri kaynağına karşı çalıştırılacak kodu sağlar */
            //        OleDbCommand command = new OleDbCommand("SELECT * FROM [Sehir$]", connection);
            //        return Execute(command);
            //    }
            //    else if (type == "Kisi")
            //    {
            //        OleDbCommand command = new OleDbCommand("SELECT * FROM [Kisi$]", connection);
            //        return Execute(command);
            //    }
            //    else if (type == "Ilce")
            //    {
            //        OleDbCommand command = new OleDbCommand("SELECT * FROM [Ilce$]", connection);
            //        return Execute(command);
            //    }
            //    else return null;
            //}
            //catch
            //{
            //    throw;
            //}
        }

        //INSERT
        public DataTable Insert<T> (T newEntity) where T : Entity
        {
            connection.ConnectionString = connectionString;
            if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                connection.Open();
            
            System.Reflection.PropertyInfo[] properties = newEntity.GetType().GetProperties();
            string propertyNames = "";
            string propertyValues = "";

            foreach (System.Reflection.PropertyInfo item in properties) {
                propertyNames += item.Name + ", ";
                propertyValues += item.GetValue(newEntity, null) + "', '";
                //OleDbCommand command = new OleDbCommand("INSERT INTO ["+newEntity.GetType().FullName.Split('.').LastOrDefault()+"$] ("+item.Name+") VALUES ('"+item.GetValue(newEntity,null)+"')", connection);
                //Execute(command);
            }
            propertyNames = propertyNames.Substring(0, propertyNames.Length - 2); //INSERT edilecek kolonların isimlerini düzgün formata çevirmek için
            propertyValues = propertyValues.Substring(0, propertyValues.Length - 3);
            OleDbCommand command = new OleDbCommand("INSERT INTO [" + newEntity.GetType().FullName.Split('.').LastOrDefault() + "$] (" + propertyNames + ") VALUES ('" + propertyValues + ")", connection);
            return Execute(command);
            //try                                                                                                                 
            //{
            //    if (newEntity is Sehir) //Gönderilen parametre bir Şehirse
            //    {
            //        Sehir newSehir = newEntity as Sehir;    //Entity'yi Sehir yap
            //        OleDbCommand command = new OleDbCommand("INSERT INTO [Sehir$] (id, ad) VALUES ('" + newSehir.id + "', '" + newSehir.ad + "')", connection);
            //        return Execute(command);
            //    }
            //    else if (newEntity is Kisi)
            //    {
            //        Kisi newKisi = newEntity as Kisi;
            //        OleDbCommand command = new OleDbCommand("INSERT INTO [Kisi$] (id, ad, soyad, yas, adres, sehir, ilce) VALUES ('" + newKisi.id +
            //            "', '" + newKisi.ad + "', '" + newKisi.soyad + "', '" + newKisi.yas + "', '" + newKisi.adres + "', '" + newKisi.sehir + "', '" + newKisi.ilce + "')", connection);
            //        return Execute(command);
            //    }
            //    else if (newEntity is Ilce)
            //    {
            //        Ilce newIlce = newEntity as Ilce;
            //        OleDbCommand command = new OleDbCommand("INSERT INTO [Ilce$] (id, sehirId, ilceAd) VALUES ('" + newIlce.id + "', '" + newIlce.sehirId + "', '" + newIlce.ad + "')", connection);
            //        return Execute(command);
            //    }
            //    else return null;
            //}
            //catch
            //{
            //    throw;
            //}
        }

        public DataTable Update<T> (T newEntity) where T : Entity
        {
            connection.ConnectionString = connectionString;
            if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                connection.Open();

            System.Reflection.PropertyInfo[] properties = newEntity.GetType().GetProperties();
            string propertyNames = "";
            string propertyValues = "";
            
            foreach (System.Reflection.PropertyInfo item in properties) {   //Propertyleri bir stringe attık
                propertyNames += item.Name + "-";
                propertyValues += item.GetValue(newEntity, null) + "-";
            }
            propertyNames = propertyNames.Substring(0, propertyNames.Length - 1);   //Sonra her propertyi bir diziye eleman olarak attık
            propertyValues = propertyValues.Substring(0, propertyValues.Length - 1);
            
            string[] propertyNamesArray = propertyNames.Split('-'); //Stringleri ayırdık
            string[] propertyValuesArray = propertyValues.Split('-');
            
            //SQL kodunun ilk ve sabit kısmı
            OleDbCommand command = new OleDbCommand("UPDATE ["+newEntity.GetType().FullName.Split('.').LastOrDefault()+"$] SET ", connection);
            
            //Her property'i düzgün formatta düzenledik
            for (int i = 0; i < propertyNamesArray.Length; i++) {
                command.CommandText += propertyNamesArray[i] + "=";
                command.CommandText += "'" + propertyValuesArray[i] + "'";
                if (i == propertyNamesArray.Length - 1) break;
                command.CommandText += ","; //Aralara virgül ekledik
            }
            command.CommandText += "WHERE " + propertyNamesArray[0] + "=" + propertyValuesArray[0];
            return Execute(command);
            //try
            //{
            //    if (newEntity is Sehir)
            //    {
            //        Sehir newSehir = newEntity as Sehir;
            //        OleDbCommand command = new OleDbCommand("UPDATE [Sehir$] SET id='" + newSehir.id + "', ad='" + newSehir.ad + "' WHERE id=" + newSehir.id + "", connection);
            //        return Execute(command);
            //    }
            //    else if (newEntity is Kisi)
            //    {
            //        Kisi newKisi = newEntity as Kisi;
            //        OleDbCommand command = new OleDbCommand("UPDATE [Kisi$] SET id='" + newKisi.id + "', ad='" + newKisi.ad + "', soyad='" + newKisi.soyad + "', yas='" + newKisi.yas + "', adres='" + newKisi.adres + "', sehir='" + newKisi.sehir + "', ilce='" + newKisi.ilce + "' WHERE id =" + newKisi.id + "", connection);
            //        return Execute(command);
            //    }
            //    else if (newEntity is Ilce)
            //    {
            //        Ilce newIlce = newEntity as Ilce;
            //        OleDbCommand command = new OleDbCommand("UPDATE [Ilce$] SET id='" + newIlce.id + "', sehirId='" + newIlce.sehirId + "', ilceAd='" + newIlce.ad + "' WHERE id=" + newIlce.id + "", connection);
            //        return Execute(command);
            //    }
            //    else return null;
            //}
            //catch
            //{
            //    throw;
            //}
        }

        public DataTable Delete (Sehir newSehir)
        {
            connection.ConnectionString = connectionString;
            if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                connection.Open();

            //OleDbCommand command = new OleDbCommand("DELETE FROM [Sehir$] WHERE id=" + newSehir.id , connection);
            OleDbCommand command = new OleDbCommand("UPDATE [Sehir$] SET id='" + 0 + "', ad='" + "" + "' WHERE id=" + newSehir.id + "", connection);
            try {
                Execute(command);
            } catch {
                throw;
            }
            return null;
        }

        public DataTable Execute (OleDbCommand command)
        {
            OleDbDataReader reader = command.ExecuteReader();
            dataTable.Load(reader);
            return dataTable;
        }
    }
}

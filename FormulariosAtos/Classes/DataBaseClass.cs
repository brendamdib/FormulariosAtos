using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;

namespace FormulariosAtos.Classes
{
    public class DataBaseClass
    {
        string ConnStr = @"Provider = Microsoft.Jet.OLEDB.4.0; Data Source = C:\Documentos ATOS\Atos.mdb; User Id = admin;Password =;Persist Security Info=False;";
        OleDbConnection MyConn;

        public void ConectaAccess()
        {
            MyConn = new OleDbConnection(ConnStr);
            MyConn.Open();
        }

        public void ExecutaQuery(string Query)
        {            
            OleDbCommand Cmd = new OleDbCommand(Query, MyConn);                       
            Cmd.ExecuteNonQuery();
        }

        public OleDbDataReader DataReader(string Query)
        {
            OleDbCommand cmd = new OleDbCommand(Query, MyConn);
            OleDbDataReader dr = cmd.ExecuteReader();
            return dr;
        }

        public void FechaConexao()
        {
            MyConn.Close();
        }
    }
}

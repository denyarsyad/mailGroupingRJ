using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Devart;
using Devart.Data.Oracle;
using System.Data;
using System.Reflection;
using System.IO;
using System.Net;

namespace mailCarArrangementSystem.Class
{
    static class CmdQry
    {
        public static OracleConnection myConnection = new OracleConnection("User Id=LMES;Password=LMES;Server=10.10.100.10;Direct=True;Sid=LMES"); /*new OracleConnection(Properties.Settings.Default.ConnectionString);*/
        public static OracleCommand myCommand;

        private static void checkConnection()
        {
            try
            {
                if (myConnection.State == ConnectionState.Closed)
                {
                    myConnection.Open();
                }
            }
            catch
            {
            }
        }

        //SELECT
        public static DataTable getData(string str)
        {
            try
            {
                checkConnection();
                OracleDataAdapter myDT = new OracleDataAdapter(str, myConnection);
                DataTable dt = new DataTable();
                myDT.Fill(dt);
                if (dt.Rows.Count > 0)
                {
                    myConnection.Close();
                }
                return dt;
            }
            catch
            {
                return null;
            }
        }

        //UPDATE OR DELETE
        public static bool executeCommand(string str)
        {
            try
            {
                checkConnection();

                myCommand = new Devart.Data.Oracle.OracleCommand(str, myConnection);
                if (myCommand.ExecuteNonQuery() == 1)
                {
                    myConnection.Close();
                    return true;
                }
                else
                {
                    myConnection.Close();
                    return false;
                }
            }
            catch
            {
                return false;
            }
        }
    }

}

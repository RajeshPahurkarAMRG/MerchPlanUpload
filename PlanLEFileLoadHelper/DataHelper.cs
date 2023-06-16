using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.Odbc;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PlanLEFileLoadHelper
{
    public static class DataHelper
    {
        //Logs are written to Sql server
        static string connetionString = ConfigurationManager.ConnectionStrings["AMRGConnection"].ToString();

        //For updating Stage table in local schema
        //static string sConnStr = ConfigurationManager.ConnectionStrings["BOAIM"].ToString();

        //For updating BOAIM database
        static string sConnStrAdm = ConfigurationManager.ConnectionStrings["BOAIMADM"].ToString();

        public static void ExecNonQuery(string sql)
        {
            SqlConnection con = new SqlConnection(connetionString);
            con.Open();
            SqlCommand cmd = new SqlCommand(sql, con);
            cmd.ExecuteNonQuery();
            con.Close();
        }
        public static void ExecNonQueryOLEDB(string sql)
        {
            OleDbConnection con2 = new OleDbConnection(sConnStrAdm);
            con2.Open();
            OleDbCommand com = new OleDbCommand(sql, con2);
            com.ExecuteNonQuery();
            con2.Close();
        }
        public static void ExecNonQueryOLEDBAdmin(string sql)
        {
            OleDbConnection con2 = new OleDbConnection(sConnStrAdm);
            con2.Open();
            OleDbCommand com = new OleDbCommand(sql, con2);
            com.ExecuteNonQuery();
            con2.Close();
        }

        internal static SqlDataReader GetDataReader(string sql)
        {
            SqlConnection con = new SqlConnection(connetionString);
            con.Open();
            SqlCommand cmd = new SqlCommand(sql, con);
            return cmd.ExecuteReader(CommandBehavior.CloseConnection);

        }

        internal static string ExecuteScalar(string sql)
        {
            string sVal = "";
            SqlConnection con = new SqlConnection(connetionString);
            con.Open();
            SqlCommand cmd = new SqlCommand(sql, con);

            using (SqlDataReader odr = cmd.ExecuteReader(CommandBehavior.CloseConnection))
            {
                while (odr.Read())
                {
                    sVal = odr[0].ToString();
                }
            }
            return sVal;

        }
        internal static string GetTableRowCount(string sTbl)
        {
            string sql = "Select count(1) from " + sTbl;
            return DataHelper.ExecuteScalar(sql);
        }

        //public static void BulkInsertDataDataTable(string sqlTableName, DataTable dataTable)
        //{
        //    using (SqlConnection connection = new SqlConnection(connetionString))
        //    {
        //        SqlBulkCopy bulkCopy =
        //            new SqlBulkCopy
        //            (
        //            connection,
        //            SqlBulkCopyOptions.TableLock |
        //            SqlBulkCopyOptions.FireTriggers |
        //            SqlBulkCopyOptions.UseInternalTransaction,
        //            null
        //            );

        //        bulkCopy.DestinationTableName = "[" + sqlTableName + "]";
        //        connection.Open();
        //        //for truncate previous data
        //        SqlCommand cmd = new SqlCommand();
        //        cmd.CommandType = CommandType.Text;
        //        cmd.CommandText = " truncate table [" + sqlTableName + "]";
        //        cmd.Connection = connection;
        //        cmd.ExecuteNonQuery();

        //        bulkCopy.BatchSize = 10000;
        //        bulkCopy.BulkCopyTimeout = 0;
        //        bulkCopy.WriteToServer(dataTable);
        //        connection.Close();
        //    }
        //}
    }
}

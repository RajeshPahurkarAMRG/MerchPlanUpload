using System;

namespace PlanLEFileLoadHelper
{
       public static class SQLLogger
        {
            private static Guid sessionid = Guid.NewGuid();

            private static string app = "LoadMerchantPlanData";
            public static string GetSessionId()
            {
                return sessionid.ToString();
            }

        public static void WriteLog(string module, string logText, int logLevel)
        {
            string sLog = logText.Replace("'", "''").Substring(0, Math.Min(logText.Length, 255));
            string sSql = "Insert into Logs(SessionId ,App ,Module ,LogText ,LogDtTm ,LogLevel) values ('"
                + sessionid
                + "','" + app + "'"
                + ",'" + module + "'"
                + ",'" + sLog + "'"
                + ",getdate(),"
                + logLevel + ")";

            DataHelper.ExecNonQuery(sSql);
        }

        public static void WriteStart()
        {
            string sLog = "Started Job";
            string module = "Main";

            string sSql = "Insert into ExecLog(SessionId ,App ,Module ,LogText,StartDtTm,ErrorDtTm ,EndDtTm ,LogDtTime) values ("
                + "'" + sessionid + "'"
                + ",'" + app + "'"
                + ",'" + module + "'"
                + ",'" + sLog + "'"
                + ",getdate(),null,null,getdate())";

            DataHelper.ExecNonQuery(sSql);

            WriteLog(module, sLog, 1);
        }

        public static void WriteExecLog(string sLog)
        {

            string module = "Main";

            string sSql = "Insert into ExecLog(SessionId ,App ,Module ,LogText,StartDtTm,ErrorDtTm ,EndDtTm ,LogDtTime) values ("
                + "'" + sessionid + "'"
                + ",'" + app + "'"
                + ",'" + module + "'"
                + ",'" + sLog + "'"
                + ",getdate(),null,null,getdate())";

            DataHelper.ExecNonQuery(sSql);
        }
        public static void WriteEnd()
        {
            string sLog = "Finished Job Successfully";
            string module = "Main";
            string sSql = "Update ExecLog Set EndDtTm=getdate() "
                + ",LogText = '" + sLog + "'"
                + " where SessionId = '" + sessionid + "'";


            DataHelper.ExecNonQuery(sSql);
            WriteLog(module, sLog, 1);
        }

        public static void WriteError(string sessionId, string ErrText, string module)
        {
            string sLog = "Error:" + ErrText.Replace("'", "''").Substring(0, 255);
            string sSql = "Update ExecLog Set ErrorDtTm=getdate(),LogText = '" + sLog + "' where SessionId = '" + sessionId + "'";
            DataHelper.ExecNonQuery(sSql);
            WriteLog(module, sLog, 1);
        }

        private static string GetLogTextFixed(string logTxt)
        {
            return logTxt.Replace("'", "''").PadLeft(255);
        }
    }

}

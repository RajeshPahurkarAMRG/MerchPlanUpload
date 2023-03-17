
using log4net;
using System;
using System.Reflection;


namespace PlanLEFileLoadHelper
{
    public static class Logger
    {
        private static readonly ILog log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

      
        public static void LogMessage(string str)
        {
            log.Info(str);
        }

        public static void LogStart()
        {
            string module = "Main";
            int loglevel = 1;
            string stext = "Job Started";
            log.Info(stext);
            SQLLogger.WriteStart();
        }

        public static void LogInsert(string Logs)
        {
            string module = "Main";
            int loglevel = 2;
            string stext = Logs;
            log.Info(stext);
            SQLLogger.WriteLog(module, Logs, loglevel);
        }

        public static void LogEnd()
        {
            string module = "Main";
            int loglevel = 1;
            string stext = "Job Finished";
            log.Info(stext);
            SQLLogger.WriteEnd();
        }

        public static void LogError(string module, Exception ex)
        {
            int loglevel = 1;
            string stext = ex.ToString();
            log.Info(stext);
            SQLLogger.WriteLog(module, ex.Message, loglevel);
        }
    }
}

using log4net;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace LoadMerchantPlanData
{
    public class Logger
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

        public static void LogExec(string Logs)
        {
            string module = "Main";
            int loglevel = 1;
            string stext = Logs;
            log.Info(stext);
            SQLLogger.WriteExecLog(Logs);
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

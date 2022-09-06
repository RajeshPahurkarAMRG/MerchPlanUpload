using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace LoadMerchantPlanData
{
    internal class Program
    {
        static string sAdminEmail = ConfigurationManager.AppSettings["AdminEmail"];
        static string sTo = ConfigurationManager.AppSettings["To"];
        static void Main(string[] args)
        {
            JobStarted();
            ProcessExcel p = new ProcessExcel();
            p.process();
            JobFinished();
        }

        private static void JobStarted()
        {
            Logger.LogStart();
            SendSMTPMail sendSMTPMail = new SendSMTPMail();
            sendSMTPMail.Dosend(sTo, "Job Started", "Started :LoadMerchantPlanData Job");
        }

        private static void JobFinished()
        {
            SendSMTPMail sendSMTPMail = new SendSMTPMail();
            sendSMTPMail.Dosend(sTo, "Job Finished", "Finished : LoadMerchantPlanData Job");
            Logger.LogEnd();
        }
    }
}

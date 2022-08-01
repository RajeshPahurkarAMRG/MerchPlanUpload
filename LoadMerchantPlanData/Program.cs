using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace LoadMerchantPlanData
{
    internal class Program
    {
        static void Main(string[] args)
        {
            logger.writelog("Started");
            ProcessExcel p = new ProcessExcel();
            p.process();
            logger.writelog("End");
        }
    }
}

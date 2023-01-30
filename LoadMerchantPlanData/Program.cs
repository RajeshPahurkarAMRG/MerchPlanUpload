
using PlanLEFileLoadHelper;
using System;
using System.Configuration;
using System.IO;

namespace LoadPlanDataBatch
{
    internal class Program
    {
        SendSMTPMail mail = new SendSMTPMail();
        static string sAdminEmail = ConfigurationManager.AppSettings["AdminEmail"];
        static string sTo = ConfigurationManager.AppSettings["To"];
        static string sArchivePath = ConfigurationManager.AppSettings["ArchiveLocalFolderDirectoryPath"].ToString();
        static string folderpath = ConfigurationManager.AppSettings["LocalFolderDirectoryPath"].ToString();
        static string sMinRowCount = ConfigurationManager.AppSettings["MinTableRowCount"].ToString();
        static string sFilePattern = "Ecom Daily Plan*.xlsx";

        static void Main(string[] args)
        {
            JobStarted();

            ValidateParams();
            int minRowCount = 1000;
            if (sMinRowCount != "")
            {
                minRowCount = int.Parse(sMinRowCount);
            }

            DirectoryInfo d = new DirectoryInfo(folderpath);
            FileInfo[] allFiles = d.GetFiles(folderpath);

            if (allFiles.Length == 0)
            {
                Logger.LogInsert("No files to process in folder " + folderpath + ". File name should be " + sFilePattern + ",exiting");
                return;
            }

            //for truncate data using loaddate
            if (allFiles.Length > 1)
            {
                Logger.LogInsert("Please process one file at a time, exiting");
                return;
            }

            if (allFiles.Length == 1)
            {
                string sFileName = allFiles[0].FullName;
                LoadPlanLE loadplanLE = new LoadPlanLE();
                loadplanLE.ProcessFile(sFileName,sArchivePath, minRowCount);
               
            }


            JobFinished();
        }
       

        private static void ValidateParams()
        {
            if (sAdminEmail == "")
                throw new Exception("Admin email is not specified");

            if (sArchivePath == "")
                throw new Exception("File Archive Path is not specified");

            if (folderpath == "")
                throw new Exception("Source folder path is not specified");

            if (!Directory.Exists(folderpath))
                throw new Exception("Source folder path not found");

            if (!Directory.Exists(sArchivePath))
                throw new Exception("Archive folder path not found");

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

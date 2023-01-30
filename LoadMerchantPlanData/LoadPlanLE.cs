using PlanLEFileLoadHelper;
using System;
using System.IO;

namespace LoadPlanDataBatch
{
    internal class LoadPlanLE
    {
        internal void ProcessFile(string sFileName, string sArchivePath, int minRowCount)
        {
            ProcessExcel p = new ProcessExcel();
            if (p.Validate(sFileName, minRowCount))
            {
                if (p.LoadData(sFileName, sArchivePath, minRowCount))
                {
                    ArchiveFile(sFileName, sArchivePath);
                }
            }

        }
        private void ArchiveFile(string fullName,string  sArchivePath)
        {
            Logger.LogInsert("Started Archiving file :" + fullName);

            string sFinalPathmoveTo = sArchivePath + AppendTimeStamp(fullName);
            File.Move(fullName, sFinalPathmoveTo);

            Logger.LogInsert("Finished Archiving file :" + fullName);
        }

        public static string AppendTimeStamp(string fileName)
        {
            return string.Concat(
                Path.GetFileNameWithoutExtension(fileName),
                DateTime.Now.ToString("yyyyMMddHHmmssfff"),
                Path.GetExtension(fileName)
                );
        }

    }
}
    

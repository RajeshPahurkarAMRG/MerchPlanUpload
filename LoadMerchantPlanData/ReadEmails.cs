using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OutLook = Microsoft.Office.Interop.Outlook;
namespace LoadMerchantPlanData
{
    internal class ReadEmails
    {

        // NOT NEEDED ANYMORE
        public void SaveAttachments()
        {
            OutLook.Application objOL;
            OutLook.Attachments objAttachments;
            OutLook.MailItem objMsg;
            OutLook.Selection objSelection;
            string oFldr;
            OutLook.Folder oFolder;
            OutLook.Folder oFolderDaily;
            int i = 0;
            int lngCount;
            string strFile;
            string strFolderPath;
            string strDeletedFiles;

            strFolderPath = "D:\\Temp\\Attachments\\";

            objOL=new OutLook.Application();
            OutLook.NameSpace oNS = objOL.GetNamespace("MAPI");
            var oMailBox = "Rajesh.Pahurkar@AMRetailGroup.com";
            oFldr = "Inbox";
            OutLook.MAPIFolder myMailBox = oNS.Folders[oMailBox];
            var omFolder = myMailBox.Folders[oFldr];
            var omFolderDaily = omFolder.Folders["UpnetEmails"];

            foreach( dynamic oMsg in omFolderDaily.Items)
            {
                if (oMsg.Class == 43)
                {
                    if (oMsg.Sender.Address == "gen2report@upnettec.com")
                    {
                        objAttachments = oMsg.Attachments;
                        lngCount = objAttachments.Count;
                        if (lngCount > 0)
                        {
                            for (i = lngCount; i > 0; i--)
                            {
                                strFile = objAttachments[i].FileName;
                                int dotIndex = strFile.LastIndexOf(".");
                                string sFileName = strFile.Substring(0, dotIndex);
                                string extn = strFile.Substring(strFile.Length - 3, 3);
                                strFile = strFolderPath + sFileName.Replace(" ", "_") + "_" + oMsg.CreationTime.ToString("MMddyyyyHHmmss") + "." + extn;
                                objAttachments[i].SaveAsFile(strFile);
                            }
                        }
                    }
                }
             
            }

        }


        //strFolderpath = CreateObject("WScript.Shell").SpecialFolders(16)



        //'Set objSelection = objOL.ActiveExplorer.Selection
        //For Each objMsg In oFolderDaily.Items
        //    If objMsg.Class = olMail Then

        //    If (objMsg.Sender = "gen2report@upnettec.com") Then

        //    Set objAttachments = objMsg.Attachments
        //    lngCount = objAttachments.Count
        //    strDeletedFiles = ""

        //        If lngCount > 0 Then

        //            For i = lngCount To 1 Step -1
        //                strFile = objAttachments.Item(i).FileName
        //                dotIndex = InStrRev(strFile, ".")
        //                sFileName = Left(strFile, dotIndex - 1)
        //                extn = Right(strFile, Len(strFile) - dotIndex)
        //                strFile = strFolderpath & Replace(sFileName, " ", "_") & "_" & Replace(Replace(objMsg.CreationTime, " ", "_"), "/", "_") & "." & extn
        //                objAttachments.Item(i).SaveAsFile strFile
        //'                objAttachments.Item(i).Delete
        //'                If objMsg.BodyFormat <> olFormatHTML Then
        //'                    strDeletedFiles = strDeletedFiles & vbCrLf & "<file://" & strFile & ">"
        //'                Else
        //'                    strDeletedFiles = strDeletedFiles & "<br>" & "<a href='file://" & _
        //'                    strFile & "'>" & strFile & "</a>"
        //'                End If
        //            Next i

        //            ' Adds the filename string to the message body and save it
        //            ' Check for HTML body
        //            If objMsg.BodyFormat<> olFormatHTML Then
        //                objMsg.Body = vbCrLf & "The file(s) were saved to " & strDeletedFiles & vbCrLf & objMsg.Body
        //            Else
        //                objMsg.HTMLBody = "<p>" & "The file(s) were saved to " & strDeletedFiles & "</p>" & objMsg.HTMLBody
        //            End If
        //            objMsg.Save
        //        End If
        //    End If
        //  End If
        //Next

        //ExitSub:

        //Set objAttachments = Nothing
        //Set objMsg = Nothing
        //Set objSelection = Nothing
        //Set objOL = Nothing
        //End Sub


    }
}

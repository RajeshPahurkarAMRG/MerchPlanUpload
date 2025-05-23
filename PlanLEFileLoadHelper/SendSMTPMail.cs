﻿using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace PlanLEFileLoadHelper
{
    public class SendSMTPMail
    {
        public string SMTPHost { get; }
        public string SMTPFromEmail { get; }
        public string SMTPFromName { get; }
        public string SMTPPort { get; }
        public string sSendEmail { get; }

        public SendSMTPMail()
        {
            SMTPHost = ConfigurationManager.AppSettings["Host"];
            SMTPFromEmail = ConfigurationManager.AppSettings["FromEmail"];
            SMTPFromName = ConfigurationManager.AppSettings["FromName"];
            SMTPPort = ConfigurationManager.AppSettings["Port"];
            sSendEmail = ConfigurationManager.AppSettings["IsMailSend"].ToString();
        }
        public void Dosend(string sTo, string sMsg, string vSubject)
        {
            if (sSendEmail.ToLower().Trim() != "true")
            {
                Logger.LogMessage("Sending email turned off");
                return;
            }
                MailMessage msg = new MailMessage();

            msg.To.Add(new MailAddress(sTo, sTo));
            msg.From = new MailAddress(SMTPFromEmail, SMTPFromName);
            msg.Subject = vSubject;
            msg.Body = sMsg;
            msg.IsBodyHtml = true;

            SmtpClient client = new SmtpClient();
            client.UseDefaultCredentials = false;
            client.Port = int.Parse(SMTPPort);
            client.Host = SMTPHost;
            client.DeliveryMethod = SmtpDeliveryMethod.Network;

            try
            {
                client.Send(msg);
                Logger.LogMessage("Message Sent Succesfully");
            }
            catch (Exception ex)
            {
                Logger.LogError("mail",ex);
            }
        }
    }

}

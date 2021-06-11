using System;
using System.Collections.Generic;
using System.Configuration;
using System.Net;
using System.Net.Mail;
using System.Text;

namespace NewmanCMD
{
    class MailModel
    {
        public static void SendMail(string strReceiverAct, string strFilePath)
        {
            //設定smtp主機
            string strSmtpAddress = "smtp.gmail.com";
            int intPort = 587;
            bool isEnableSSL = true;
            //填入寄送方email和密碼
            string strEmailFrom = ConfigurationManager.AppSettings["mailAccount"];
            string strPassword = ConfigurationManager.AppSettings["mailPassword"];
            //收信方email
            string strEmailTo = strReceiverAct;
            //主旨
            string strSubject = ConfigurationManager.AppSettings["mailSubject"];
            //內容
            string StrBody = ConfigurationManager.AppSettings["mailBody"];

            using (MailMessage mail = new MailMessage())
            {
                mail.From = new MailAddress(strEmailFrom);
                mail.To.Add(strEmailTo);
                mail.Subject = strSubject;
                mail.Body = StrBody;
                // 若你的內容是HTML格式，則為True
                mail.IsBodyHtml = false;

                //夾帶檔案
                mail.Attachments.Add(new Attachment(strFilePath));

                using (SmtpClient smtp = new SmtpClient(strSmtpAddress, intPort))
                {
                    smtp.Credentials = new NetworkCredential(strEmailFrom, strPassword);
                    smtp.EnableSsl = isEnableSSL;
                    smtp.Send(mail);
                }
            }

        }
    }
}

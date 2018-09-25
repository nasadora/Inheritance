using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.Mail;


class Mail
{
    private AppConfig _appCong = new AppConfig();

    public void SendMail(string subject, string body)
    {
        MailMessage mail = new MailMessage();
        mail.To.Add(_appCong.GetRecipients());
        mail.Subject = subject;
        mail.Body = body;

        MailAddress sender = new MailAddress(_appCong.GetMailboxAdrress(), _appCong.GetDisplaySender());
        mail.From = sender;
        mail.Sender = sender;

        SmtpClient smtp = new SmtpClient(_appCong.GetSmtpHost());
        try
        {
            smtp.Send(mail);
        }
        catch (Exception e)
        {
            Log logFile = new Log();
            logFile.LogMessageToFile("ERROR| Error occured during sending emails\n" + e.ToString() + "\n\n\n");
        }
    }


    public void SendMailWithAttachment(string subject, string body, string filePath)
    {
        MailMessage mail = new MailMessage();
        mail.To.Add(_appCong.GetRecipients());
        mail.CC.Add(_appCong.GetCC());
        mail.Subject = subject;
        mail.Body = body;

        MailAddress sender = new MailAddress(_appCong.GetMailboxAdrress(), _appCong.GetDisplaySender());
        mail.From = sender;
        mail.Sender = sender;


        using (Attachment attachment = new Attachment(filePath))
        {
            mail.Attachments.Add(attachment);

            SmtpClient smtp = new SmtpClient(_appCong.GetSmtpHost());
            try
            {
                smtp.Send(mail);
            }
            catch (Exception e)
            {
                Log logFile = new Log();
                logFile.LogMessageToFile("ERROR| Error occured during sending emails\n" + e.ToString() + "\n\n\n");
            }
        }
    }

    public void SendMail(string to, string subject, string body)
    {
        MailMessage mail = new MailMessage();
        mail.To.Add(to);
        mail.Subject = subject;
        mail.Body = body;

        MailAddress sender = new MailAddress(_appCong.GetMailboxAdrress(), _appCong.GetDisplaySender());
        mail.From = sender;
        mail.Sender = sender;

        SmtpClient smtp = new SmtpClient(_appCong.GetSmtpHost());
        try
        {
            smtp.Send(mail);
        }
        catch (Exception e)
        {
            Log logFile = new Log();
            logFile.LogMessageToFile("ERROR - Error occured during sending emails\n" + e.ToString() + "\n\n\n");
        }
    }


    public static MailAddressCollection ParseMailsListToMailsCollection(string str)
    {
        MailAddressCollection list = new MailAddressCollection();
        string[] mails = str.Split(new char[] { ',', ';' });
        foreach (string mail in mails)
        {
            list.Add(mail);

        }
        return list;
    }

}


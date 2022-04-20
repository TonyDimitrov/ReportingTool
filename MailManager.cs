using ReportExtraction.DTO;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;
using Serilog;

namespace ReportExtraction
{
    public class MailManager : ReportManager
    {
        private readonly EmailParameters parameters;

        public MailManager()
        {
            this.parameters = new EmailParameters();
        }

        public void SendAttachedReport(ExtractionReportDto reportParams, Stream contentAsStream, string fileName)
        {
            var mail = new MailMessage
            {
                From = new MailAddress(parameters.NoReplyEmail),

                Subject = reportParams.EmailSubject,
                IsBodyHtml = false,
                Body = reportParams.EmailText
            };

            var recipients = reportParams.Email
                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(e => e.Trim());

            AddRecipients(mail, recipients);

            AddAttachment(mail, contentAsStream, fileName);
            SendEmail(mail);
        }

        public void SendAttachedReport(ExtractionReportDto reportParams, string fileContent, string fileName)
        {
            Stream contentAsStream;
            using (contentAsStream = new MemoryStream(Encoding.UTF8.GetBytes(fileContent ?? "")))
            {
                SendAttachedReport(reportParams, contentAsStream, fileName);
            }
        }

        public void SendZipAttachedReport(ExtractionReportDto reportParams, byte[] bytesContent, string fileName)
        {
            var mail = new MailMessage
            {
                From = new MailAddress(parameters.NoReplyEmail),

                Subject = reportParams.EmailSubject,
                IsBodyHtml = false,
                Body = reportParams.EmailText
            };

            var recipients = reportParams.Email
                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(e => e.Trim());

            AddRecipients(mail, recipients);

            var zipPath = CreateZipReport(bytesContent, fileName, reportParams.ExtractionId);

            mail.Attachments.Add(new Attachment(zipPath, MediaTypeNames.Application.Zip));

            SendEmail(mail);
        }

        public void SendZipAttachedReport(ExtractionReportDto reportParams, string text, string fileName)
        {
            var mail = new MailMessage
            {
                From = new MailAddress(parameters.NoReplyEmail),

                Subject = reportParams.EmailSubject,
                IsBodyHtml = false,
                Body = reportParams.EmailText
            };

            var recipients = reportParams.Email
                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(e => e.Trim());

            AddRecipients(mail, recipients);

            var zipPath = CreateZipReport(text, fileName, reportParams.ExtractionId);

            mail.Attachments.Add(new Attachment(zipPath, MediaTypeNames.Application.Zip));

            SendEmail(mail);
        }

        public void SendError(string errorRecipients, string errorText)
        {
            var mail = new MailMessage
            {
                From = new MailAddress(parameters.NoReplyEmail),

                Subject = $"Error while processing report! {DateTime.UtcNow.ToString("dd/MM/yyyy")}",
                IsBodyHtml = false,
                Body = errorText
            };

            var recipients = errorRecipients
                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(e => e.Trim());

            AddRecipients(mail, recipients);
            SendEmail(mail);
        }

        private void AddRecipients(MailMessage mail, IEnumerable<string> team)
        {
            foreach (var to in team)
            {
                mail.To.Add(new MailAddress(to));
            }
        }

        private void AddAttachment(MailMessage mail, Stream stream, string fileName)
        {
            var attachment = new Attachment(stream, fileName);

            mail.Attachments.Add(attachment);
        }

        private void SendEmail(MailMessage mail)
        {
            try
            {
                var smtp = new SmtpClient(parameters.SmtpClient);
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.UseDefaultCredentials = false;
                smtp.EnableSsl = true;
                smtp.Port = parameters.EmailPort;

                smtp.Credentials = new System.Net.NetworkCredential(parameters.EmailUsername, parameters.EmailPassword);

                try
                {
                    smtp.Send(mail);
                }
                catch (Exception ex)
                {
                    Log.Error(ex, $"Error on sending email with: {parameters.EmailUsername}!");
                    try
                    {
                        smtp.Credentials = new System.Net.NetworkCredential("", "");
                        smtp.Send(mail);
                    }
                    catch (Exception ex2)
                    {
                        Log.Error(ex2, $"Error on sending email with: !");
                        try
                        {
                            smtp.Credentials = new System.Net.NetworkCredential("", "");
                            smtp.Send(mail);
                        }
                        catch (Exception ex3)
                        {
                            Log.Error(ex3, $"Error on sending email with: !");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Error on sending email!");
                throw;
            }
        }    
    }
}

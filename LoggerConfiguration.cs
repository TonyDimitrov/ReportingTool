using Serilog;
using System;
using System.Configuration;

namespace ReportExtraction
{
    public class LoggerConfiguration
    {
        public LoggerConfiguration()
        {
            Log.Logger = new Serilog.LoggerConfiguration()
                 .ReadFrom.AppSettings()
                 .CreateLogger();
        }

        public void LogError(Exception ex, string errorMsg, bool sendMail)
        {
            Log.Error(ex, errorMsg);

            if (sendMail)
            {
                var recipient = ConfigurationManager.AppSettings.Get("error_mail_recipient");
                new MailManager().SendError(recipient, errorMsg);
            }
        }

        public void LogWarning(string warningMsg, bool sendMail)
        {
            Log.Warning(warningMsg);

            if (sendMail)
            {
                var recipient = ConfigurationManager.AppSettings.Get("error_mail_recipient");
                new MailManager().SendError(recipient, warningMsg);
            }
        }
    }
}

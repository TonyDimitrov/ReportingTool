using System.Configuration;

namespace ReportExtraction.DTO
{
   public class EmailParameters : ConfigurationSection
    {
        public EmailParameters()
        {
            this.EmailPort = int.Parse(ConfigurationManager.AppSettings.Get("port"));
            this.SmtpClient = ConfigurationManager.AppSettings.Get("smtpClient");
            this.EmailFrom = ConfigurationManager.AppSettings.Get("emailFrom");
            this.EmailUsername = ConfigurationManager.AppSettings.Get("emailUsername");
            this.EmailPassword = ConfigurationManager.AppSettings.Get("emailPassword");
            this.NoReplyEmail = ConfigurationManager.AppSettings.Get("noReplyEmail");
        }

        public string NoReplyEmail { get; set; }
        public int EmailPort { get; set; }
        public string SmtpClient { get; set; }
        public string EmailFrom { get; set; }
        public string EmailUsername { get; set; }
        public string EmailPassword { get; set; }
    }
}
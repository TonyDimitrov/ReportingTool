namespace ReportExtraction.DTO
{
    public class FtpCredentials
    {
        public ushort Port { set; get; }
        public string FtpHost { set; get; }
        public string Username { set; get; }
        public string Password { set; get; }
        public string DestinationFolderPath { set; get; }
        public string DownloadPath { set; get; } = "IN";
        public string UploadPath { set; get; } = "OUT";
    }
}

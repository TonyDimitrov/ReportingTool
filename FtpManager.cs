using ReportExtraction.DTO;
using Serilog;
using System;
using System.IO;
using System.Net;

namespace ReportExtraction
{
    public class FtpManager : ReportManager
    {
        public override string FullPathReportFileDirectory => Path.Combine(SpExportReportDirectory, ReportFtpDirectory);

        public override string FullPathZipReportFileDirectory => Path.Combine(SpExportReportDirectory, ZipReportFtpDirectory);

        public void SendReport(FtpCredentials ftpData, string filename, byte[] fileContent)
        {
            try
            {
                var ftp = CreateFtpRequest(ftpData, filename);

                Stream ftpstream = ftp.GetRequestStream();
                ftpstream.Write(fileContent, 0, fileContent.Length);
                ftpstream.Close();
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Error sending file via FTP!");
            }
        }

        public void SendZipReport(FtpCredentials ftpData, string fileName, byte[] fileContent, string extractionId)
        {
            try
            {
                var ftp = CreateFtpRequest(ftpData, ZipReportFtpDirectory.AddDirectoryIdentifier(extractionId));

                var zipPath = CreateZipReport(fileContent, fileName, extractionId);
                var zipContent = File.ReadAllBytes(zipPath);
                Stream ftpstream = ftp.GetRequestStream();

                ftpstream.Write(zipContent, 0, zipContent.Length);
                ftpstream.Close();
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Error sending file via FTP!");
            }
        }

        private FtpWebRequest CreateFtpRequest(FtpCredentials ftpData, string filename)
        {
            var ftp = (FtpWebRequest)WebRequest.Create(ftpData.FtpHost + "/" + ftpData.DestinationFolderPath + "/" + filename);

            ftp.AuthenticationLevel = System.Net.Security.AuthenticationLevel.None;
            ftp.Credentials = new NetworkCredential(ftpData.Username, ftpData.Password);
            ftp.KeepAlive = true;
            ftp.UseBinary = true;
            ftp.Method = WebRequestMethods.Ftp.UploadFile;

            return ftp;
        }
    }
}

using System;
using System.Configuration;
using System.IO;
using System.IO.Compression;

namespace ReportExtraction
{
    public abstract class ReportManager
    {
        protected readonly string SpExportReportDirectory;

        protected readonly string ReportMailDirectory;
        protected readonly string ZipReportMailDirectory;

        protected readonly string ReportFtpDirectory;
        protected readonly string ZipReportFtpDirectory;

        protected readonly LoggerConfiguration logger = new LoggerConfiguration();

        public ReportManager()
        {
            this.SpExportReportDirectory = ConfigurationManager.AppSettings.Get("spExportReportDirectory");

            this.ReportMailDirectory = ConfigurationManager.AppSettings.Get("reportMailDirectory");
            this.ZipReportMailDirectory = ConfigurationManager.AppSettings.Get("zipReportMailDirectory");

            this.ReportFtpDirectory = ConfigurationManager.AppSettings.Get("reportFtpDirectory");
            this.ZipReportFtpDirectory = ConfigurationManager.AppSettings.Get("zipReportFtpDirectory");
        }

        public virtual string FullPathReportFileDirectory => Path.Combine(SpExportReportDirectory, ReportMailDirectory);

        public virtual string FullPathZipReportFileDirectory => Path.Combine(SpExportReportDirectory, ZipReportMailDirectory);

        public virtual string CreateZipReport(string fileContent, string fileName, string extractionId)
        {
            try
            {
                NewZipDirectory(extractionId);

                File.WriteAllText(Path.Combine(FullPathReportFileDirectory.AddDirectoryIdentifier(extractionId), fileName), fileContent);

                if (File.Exists(FullPathZipReportFileDirectory.AddDirectoryIdentifier(extractionId)))
                {
                    File.Delete(FullPathZipReportFileDirectory.AddDirectoryIdentifier(extractionId));
                }

                ZipFile.CreateFromDirectory(FullPathReportFileDirectory.AddDirectoryIdentifier(extractionId),
                    FullPathZipReportFileDirectory.AddDirectoryIdentifier(extractionId));

                return FullPathZipReportFileDirectory.AddDirectoryIdentifier(extractionId);
            }
            catch (Exception ex)
            {
                logger.LogError(ex, $"Error on creating ZIP file! File name: [{fileName}], Extraction ID: [{extractionId}] {Environment.NewLine}" +
                    $"Exception: [{ex.Message}]", true);
                throw;
            }
        }

        public virtual string CreateZipReport(byte[] bytesContent, string fileName, string extractionId)
        {
            try
            {
                NewZipDirectory(extractionId);

                File.WriteAllBytes(Path.Combine(FullPathReportFileDirectory.AddDirectoryIdentifier(extractionId), fileName), bytesContent);

                if (File.Exists(FullPathZipReportFileDirectory.AddDirectoryIdentifier(extractionId)))
                {
                    File.Delete(FullPathZipReportFileDirectory.AddDirectoryIdentifier(extractionId));
                }
                ZipFile.CreateFromDirectory(FullPathReportFileDirectory.AddDirectoryIdentifier(extractionId), FullPathZipReportFileDirectory.AddDirectoryIdentifier(extractionId));

                return FullPathZipReportFileDirectory.AddDirectoryIdentifier(extractionId);
            }
            catch (Exception ex)
            {
                logger.LogError(ex, $"Error on creating ZIP file! File name: [{fileName}], Extraction ID: [{extractionId}] {Environment.NewLine}" +
                 $"Exception: [{ex.Message}]", true);
                throw;
            }
        }

        private void NewZipDirectory(string extractionId)
        {
            if (!Directory.Exists(SpExportReportDirectory))
            {
                Directory.CreateDirectory(SpExportReportDirectory);
            }

            if (!Directory.Exists(FullPathReportFileDirectory.AddDirectoryIdentifier(extractionId)))
            {
                Directory.CreateDirectory(FullPathReportFileDirectory.AddDirectoryIdentifier(extractionId));
            }
            else
            {
                Directory.Delete(FullPathReportFileDirectory.AddDirectoryIdentifier(extractionId), true);
                Directory.CreateDirectory(FullPathReportFileDirectory.AddDirectoryIdentifier(extractionId));
            }
        }
    }
}

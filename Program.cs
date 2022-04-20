using ReportExtraction.DTO;
using ReportExtraction.DTO.Enums;
using System.IO;
using System.Text;

namespace ReportExtraction
{
    class Program
    {
        private static readonly LoggerConfiguration logger = new LoggerConfiguration();

        static void Main(string[] args)
        {
            var da = new DataAccessor();

            var reportParams = da.GetReportParameters()
                .GetAwaiter()
                .GetResult();

            var filter = new ReportFilters();
            reportParams = filter.ApplySendReportFilters(reportParams);

            while (reportParams.Count > 0)
            {
                var parameters = reportParams.Dequeue();

                if (parameters.FileFormat == FileFormatType.Csv)
                {
                    var result = da.CsvFormatBuilder(parameters.StoreProcedure)
                        .GetAwaiter()
                        .GetResult();

                    if (ProtocolType.Mail == parameters.ProtocolType || ProtocolType.Other == parameters.ProtocolType)
                    {
                        ProcessCsvExportByMail(parameters, result);
                    }
                    else if (ProtocolType.Ftp == parameters.ProtocolType)
                    {
                        ProcessCsvExportByFtp(parameters, result);
                    }
                    else if (ProtocolType.MailAndFtp == parameters.ProtocolType)
                    {
                        ProcessCsvExportByMail(parameters, result);

                        ProcessCsvExportByFtp(parameters, result);
                    }
                }
                else if (parameters.FileFormat == FileFormatType.Excel)
                {
                    var result = da.ExcelFormatBuilder(parameters.StoreProcedure, parameters.FileName, parameters.TypeOfExcel)
                       .GetAwaiter()
                       .GetResult();

                    if (ProtocolType.Mail == parameters.ProtocolType || ProtocolType.Other == parameters.ProtocolType)
                    {
                        ProcessExcelExportByMail(parameters, result);
                    }
                    else if (ProtocolType.Ftp == parameters.ProtocolType)
                    {
                        ProcessExcelExportByFtp(parameters, result);
                    }
                    else if (ProtocolType.MailAndFtp == parameters.ProtocolType)
                    {
                        ProcessExcelExportByMail(parameters, result);

                        ProcessExcelExportByFtp(parameters, result);
                    }
                }
                else
                {
                    logger.LogWarning($"Unrecognise file format type! File name: [{parameters.FileName}], " +
                        $"File format: [{parameters.FileFormat}], report ID: [{parameters.ExtractionId}]", false);
                }
            }
        }

        private static void ProcessCsvExportByMail(ExtractionReportDto reportParams, ResultAsCsv result)
        {
            if (!result.HasResults && !reportParams.SendEmpty)
            {
                return;
            }

            var sendMail = new MailManager();

            if (reportParams.SendCompressed)
            {
                sendMail.SendZipAttachedReport(reportParams, result.TextAsCsv, reportParams.FileName);
            }
            else
            {
                sendMail.SendAttachedReport(reportParams, result.TextAsCsv, reportParams.FileName);
            }
        }

        private static void ProcessExcelExportByMail(ExtractionReportDto reportParams, ResultAsExcel result)
        {
            if (!result.HasResults && !reportParams.SendEmpty)
            {
                return;
            }

            var sendMail = new MailManager();

            if (!result.HasResults && !reportParams.SendEmpty)
            {
                return;
            }

            if (reportParams.SendCompressed)
            {
                sendMail.SendZipAttachedReport(reportParams, result.BytesAsExcel, reportParams.FileName);
            }
            else
            {
                sendMail.SendAttachedReport(reportParams, new MemoryStream(result.BytesAsExcel), reportParams.FileName);
            }
        }

        private static void ProcessCsvExportByFtp(ExtractionReportDto reportParams, ResultAsCsv result)
        {
            if (!result.HasResults && !reportParams.SendEmpty)
            {
                return;
            }

            var ftpSender = new FtpManager();

            var asBytes = Encoding.UTF8.GetBytes(result.TextAsCsv);

            if (reportParams.SendCompressed)
            {
                ftpSender.SendZipReport(reportParams.FtpCredentials, reportParams.FileName, asBytes, reportParams.ExtractionId);
            }
            else
            {
                ftpSender.SendReport(reportParams.FtpCredentials, reportParams.FileName, asBytes);
            }
        }

        private static void ProcessExcelExportByFtp(ExtractionReportDto reportParams, ResultAsExcel result)
        {

            if (!result.HasResults && !reportParams.SendEmpty)
            {
                return;
            }

            var ftpSender = new FtpManager();

            if (reportParams.SendCompressed)
            {
                ftpSender.SendZipReport(reportParams.FtpCredentials, reportParams.FileName, result.BytesAsExcel, reportParams.ExtractionId);
            }
            else
            {
                ftpSender.SendReport(reportParams.FtpCredentials, reportParams.FileName, result.BytesAsExcel);
            }
        }
    }
}

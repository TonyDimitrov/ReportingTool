using ReportExtraction.DTO.Enums;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ReportExtraction.DTO
{
    public class ExtractionReportDto
    {
        private string emailSubject;
        private string emailText;

        public ExtractionReportDto()
        {
            this.FtpCredentials = new FtpCredentials();
        }

        public string ExtractionId { get; set; }

        public string StoreProcedure { get; set; }

        public string Email { get; set; }

        public string EmailSubject
        {
            get
            {
                return this.emailSubject?.Replace("{DATE}", DateTime.UtcNow.ToString("dd/MM/yyyy"));
            }
            set
            {
                this.emailSubject = value;
            }
        }

        public FileFormatType FileFormat { get; set; }

        public string FileName { get; set; }

        public ProtocolType ProtocolType { get; set; }

        public int? TypeOfExcel { get; set; }

        public string EmailText
        {
            get
            {
                return this.emailText?.Replace("{DATE}", DateTime.UtcNow.ToString("dd/MM/yyyy"));
            }
            set
            {
                this.emailText = value;
            }
        }

        public bool SendCompressed { get; set; } = false;

        public bool SendEmpty { get; set; } = true;

        public bool? SendOnBusinessDays { get; set; }

        public string SendOnWeekDays { get; set; }

        public bool? SendOnMonthBeginning { get; set; }

        public bool? SendOnMonthEnd { get; set; }

        public DateTime? IgnoreFromDate { get; set; }

        public DateTime? IgnoreToDate { get; set; }

        public FtpCredentials FtpCredentials { get; set; }

        public IEnumerable<DayOfWeek> ParseToSendWeekDays()
        {
            var weekDays = this.SendOnWeekDays
                ?.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(x =>
                {
                    if (Enum.TryParse<WeekDay>(x.Trim(), true, out var weekDay))
                    {
                        return weekDay;
                    }

                    return WeekDay.Other;
                });

            if (weekDays == null)
            {
                weekDays = new List<WeekDay>();
            }

            foreach (var day in weekDays)
            {
                switch (day)
                {
                    case WeekDay.D1:
                        yield return DayOfWeek.Monday;
                        break;

                    case WeekDay.D2:
                        yield return DayOfWeek.Tuesday;
                        break;

                    case WeekDay.D3:
                        yield return DayOfWeek.Wednesday;
                        break;

                    case WeekDay.D4:
                        yield return DayOfWeek.Thursday;
                        break;

                    case WeekDay.D5:
                        yield return DayOfWeek.Friday;
                        break;

                    case WeekDay.D6:
                        yield return DayOfWeek.Saturday;
                        break;

                    case WeekDay.D7:
                        yield return DayOfWeek.Sunday;
                        break;
                }
            }
        }
    }
}
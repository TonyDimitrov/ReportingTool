using ReportExtraction.DTO;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ReportExtraction
{
    public class ReportFilters
    {
        public Queue<ExtractionReportDto> ApplySendReportFilters(Queue<ExtractionReportDto> reports)
        {

            Dictionary<string, ExtractionReportDto> reportsInDictionary = new Dictionary<string, ExtractionReportDto>();

            foreach (var report in reports)
            {
                var daysOfWeekSend = report.ParseToSendWeekDays();

                var hasFilters = (report.SendOnBusinessDays.HasValue && report.SendOnBusinessDays.Value)
                    || daysOfWeekSend.Any()
                    || (report.SendOnMonthBeginning.HasValue && report.SendOnMonthBeginning.Value)
                    || (report.SendOnMonthEnd.HasValue && report.SendOnMonthEnd.Value);

                var currentDate = DateTime.UtcNow.Date;

                if (hasFilters)
                {
                    if (report.SendOnBusinessDays.HasValue
                        && report.SendOnBusinessDays.Value
                        && currentDate.DayOfWeek != DayOfWeek.Saturday
                        && currentDate.DayOfWeek != DayOfWeek.Sunday)
                    {
                        reportsInDictionary[report.ExtractionId] = report;
                    }

                    // Send days of week filter
                    if (daysOfWeekSend.Count() != 0 && daysOfWeekSend.Contains(currentDate.DayOfWeek))
                    {
                        reportsInDictionary[report.ExtractionId] = report;
                    }

                    // Send beginning of month filter
                    if (report.SendOnMonthBeginning.HasValue
                        && report.SendOnMonthBeginning.Value
                        && currentDate.Day == 1)
                    {
                        reportsInDictionary[report.ExtractionId] = report;
                    }

                    // Send end of month filter
                    var lastDayOfMonth = DateTime.DaysInMonth(currentDate.Year, currentDate.Month);
                    if (report.SendOnMonthEnd.HasValue
                        && report.SendOnMonthEnd.Value
                        && currentDate.Day == lastDayOfMonth)
                    {
                        reportsInDictionary[report.ExtractionId] = report;
                    }
                }
                else
                {
                    reportsInDictionary[report.ExtractionId] = report;
                }

                // Between two dates ignore filter
                if (report.IgnoreFromDate.HasValue && report.IgnoreToDate.HasValue && report.IgnoreFromDate <= report.IgnoreToDate)
                {
                    if (report.IgnoreFromDate <= currentDate.Date && report.IgnoreToDate >= currentDate.Date)
                    {
                        if (reportsInDictionary.ContainsKey(report.ExtractionId))
                        {
                            reportsInDictionary.Remove(report.ExtractionId);
                        }
                    }
                }
            }

            return new Queue<ExtractionReportDto>(reportsInDictionary.Select(r => r.Value));
        }
    }
}

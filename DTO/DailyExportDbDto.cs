namespace ReportExtraction.DTO
{
   public class DailyExportDbDto : DailyExportBaseDto
    {
        public int IdCarrier { get; set; }
        public string FromZip { get; set; }
        public string ToZip { get; set; }
        public double? MultipacketWeight { get; set; }
        public double? MultipacketDepth { get; set; }
        public double? MultipacketHeight { get; set; }
        public double? MultipacketWidth { get; set; }
        public int? IdItalianCityFrom { get; set; }
        public int? IdItalianCityTo { get; set; }
        public int? IdCountry { get; set; }
    }
}
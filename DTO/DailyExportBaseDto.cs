
namespace ReportExtraction.DTO
{
    public class DailyExportBaseDto
    {
        public int IdUserCityOrder { get; set; }
        public int IdUser { get; set; }
        public string Login { get; set; }
        public double OrderPrice { get; set; }
        public double Sales { get; set; }
        public string OrderDate { get; set; }
        public string FromProvince { get; set; }
        public string ToCountry { get; set; }
        public string ToRegion { get; set; }
        public float Weight { get; set; }
        public string WeightRange { get; set; }
        public string WeightUpperRange { get; set; }
        public string Agent { get; set; }
        public string BuisnessName { get; set; }
        public string Status { get; set; }
        public string TargetCountry { get; set; }
        public string OrderMonth { get; set; }  
        public string OrderWeek { get; set; }  
        public string PaymentMethod { get; set; }
        public string EstimatedCost { get; set; }
        public string Carrier { get; set; }
        public string CustomerTracking { get; set; }
        public string PrivateCompany { get; set; }
        public string CouponName { get; set; }
    }
}

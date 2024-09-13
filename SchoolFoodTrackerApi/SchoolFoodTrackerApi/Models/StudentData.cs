namespace SchoolFoodTrackerApi.Models
{
    public class StudentData
    {
        public required string Klase { get; set; }
        public required string Vards { get; set; }
        public required string Uzvards { get; set; }
        public required string Kods { get; set; }
        public string? Ligums { get; set; }
        public string? IrSamaksats { get; set; }
    }

}

using System;
using System.ComponentModel.DataAnnotations;

namespace SchoolFoodTrackerApi.Models
{
    public class StudentDataRequest
    {
        [DataType(DataType.DateTime)]
        public DateTime? StartDate { get; set; }
        public DateTime? EndDate { get; set; }
        public string? Class { get; set; }
        public string? Name { get; set; }
        public string? Surname { get; set; }
        public string? Code { get; set; }
        public bool? HasContract { get; set; }
        public bool? IsPaid { get; set; }
    }
}

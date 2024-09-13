namespace SchoolFoodTrackerApi.Models
{
    public class PersonSearchModel
    {
        public class PersonInfo
        {
            public string Name { get; set; }
            public string Surname { get; set; }

            public string Class { get; set; }
            public bool HasContract { get; set; }
            public bool IsPaid { get; set; }
        }

        public class SearchResult
        {
            public string Status { get; set; }
            public PersonInfo PersonInfo { get; set; }
        }
    }
}

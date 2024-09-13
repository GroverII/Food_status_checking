using Google.Apis.Sheets.v4;
using SchoolFoodTrackerApi.Models;
using static SchoolFoodTrackerApi.Models.PersonSearchModel;

namespace SchoolFoodTrackerApi.Services
{
    public interface IGoogleSheetsService
    {
        SheetsService InitializeGoogleSheetsService();
        SearchResult FindAndPrintRow(string sheetName, int code);
        List<IList<object>> RetrieveStudentDataFromSheets(StudentDataRequest request, int month);
        List<string> GetAllClasses();
        void UpdateData(DateTime startDate, DateTime endDate);
        void UpdateHeaders(int startMonth, int endMonth);
    }
}
using SchoolFoodTrackerApi.Models;
using System.Collections.Generic;

namespace SchoolFoodTrackerApi.Services
{
    public interface ISchoolFoodService
    {
        IEnumerable<StudentData> ReadCsvFile(string filePath);
        bool IsExcludedClass(string klase);
        IEnumerable<string> GetExcelFilePaths();
        byte[] CreateExcelFile(List<object> result);
        bool AreStringsEqual(string str1, string str2);
        string ReplaceLatvianCharacters(string input);
        string ConvertIndexToColumnName(int columnIndex);
        IList<int> CalculateWorkingDaysInSpecifiedRange(StudentDataRequest request, int month, int daysInMonth);
        int CalculateStartDay(IList<int> workingDays, IList<int> monthWorkingDays);
        bool IsMonthInRange(StudentDataRequest request, int month);
        List<IList<object>> FilterValues(List<IList<object>> combinedValues, StudentDataRequest request);
        int Compare(string? x, string? y);
    }
}

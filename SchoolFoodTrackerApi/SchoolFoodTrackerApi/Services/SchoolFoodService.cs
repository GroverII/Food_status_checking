using CsvHelper;
using CsvHelper.Configuration;
using OfficeOpenXml;
using SchoolFoodTrackerApi.Models;
using System.Globalization;

namespace SchoolFoodTrackerApi.Services
{
    public class SchoolFoodService : ISchoolFoodService
    {
        public IEnumerable<StudentData> ReadCsvFile(string filePath)
        {
            var studentDataList = new List<StudentData>();

            using (var reader = new StreamReader(filePath))
            using (var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture)))
            {
                // Пропускаем первые две строки (заголовки)
                csv.Read();
                csv.Read();

                while (csv.Read())
                {
                    var excelStudentData = new StudentData
                    {
                        Klase = csv.GetField<string>(3), // Индекс 3 соответствует "Отдел №"
                        Vards = csv.GetField<string>(1), // Индекс 1 соответствует "Имя"
                        Uzvards = csv.GetField<string>(2), // Индекс 2 соответствует "Фамилия"
                        Kods = csv.GetField<string>(0), // Индекс 0 соответствует "ID сотрудника"
                        Ligums = null,
                        IrSamaksats = null
                    };

                    studentDataList.Add(excelStudentData);
                }
            }

            return studentDataList;
        }

        public bool IsExcludedClass(string klase)
        {
            string[] excludedClasses = ["skolotajs", "dzvsk", "TehPersonal", "viesiskolot", "virtuve1eka", "virtuve2eka"];
            return excludedClasses.Contains(klase, StringComparer.OrdinalIgnoreCase);
        }

        public IEnumerable<string> GetExcelFilePaths()
        {
            string excelDirectory = @"C:\Users\grigo\OneDrive\Документы\GitHub\Edinashana-sistema";

            // Поиск всех файлов Excel в указанной директории с расширением ".xlsx"
            var excelFiles = Directory.GetFiles(excelDirectory, "*.csv");

            return excelFiles;
        }

        public byte[] CreateExcelFile(List<object> result)
        {
            // Создание нового файла Excel в памяти
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var memoryStream = new MemoryStream();
            using var package = new ExcelPackage(memoryStream);

            foreach (var monthData in result)
            {
                string? monthName = (string?)monthData?.GetType()?.GetProperty("Month")?.GetValue(monthData);

                // Преобразование "Data" в IList<IList<object>>
                var monthDataList = (IList<IList<object>>?)monthData?.GetType()?.GetProperty("Data")?.GetValue(monthData);

                if (monthName != null && monthDataList != null)
                {
                    // Создание нового листа с именем месяца
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(monthName);

                    // Заполнение листа данными
                    for (int row = 1; row <= monthDataList.Count; row++)
                    {
                        var rowData = monthDataList[row - 1];
                        for (int col = 1; col <= rowData.Count; col++)
                        {
                            worksheet.Cells[row, col].Value = rowData[col - 1];
                        }
                    }
                }
            }

            // Сохранение файла Excel в памяти
            package.Save();

            // Возвращение массива байтов, представляющего Excel-файл
            return memoryStream.ToArray();
        }

        public bool AreStringsEqual(string str1, string str2)
        {
            // Заменяем латышские символы
            string normalizedStr1 = ReplaceLatvianCharacters(str1);
            string normalizedStr2 = ReplaceLatvianCharacters(str2);

            // Используем инвариантную культуру для обеспечения портируемости
            CultureInfo invariantCulture = CultureInfo.InvariantCulture;

            // Получаем объект CompareInfo для текущей культуры
            CompareInfo compareInfo = invariantCulture.CompareInfo;

            // Используем метод Compare с параметром IgnoreNonSpace
            return compareInfo.Compare(normalizedStr1, normalizedStr2, CompareOptions.IgnoreNonSpace | CompareOptions.IgnoreCase) == 0;
        }

        public string ReplaceLatvianCharacters(string input)
        {
            // Заменяем латышские символы
            return input
                .Replace("ā", "a")
                .Replace("č", "c")
                .Replace("ē", "e")
                .Replace("ģ", "g")
                .Replace("ī", "i")
                .Replace("ķ", "k")
                .Replace("ļ", "l")
                .Replace("ņ", "n")
                .Replace("š", "s")
                .Replace("ū", "u")
                .Replace("ž", "z");
        }

        public string ConvertIndexToColumnName(int columnIndex)
        {
            const string alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            return columnIndex < alphabet.Length
                ? alphabet[columnIndex].ToString()
                : alphabet[columnIndex / alphabet.Length - 1].ToString() + alphabet[columnIndex % alphabet.Length];
        }

        public IList<int> CalculateWorkingDaysInSpecifiedRange(StudentDataRequest request, int month, int daysInMonth)
        {
            var monthWorkingDays = new List<int>();

            if (request.StartDate.HasValue)
            {
                int endDay = request.EndDate?.Month == month ? request.EndDate?.Day ?? daysInMonth : daysInMonth;

                for (int i = request.StartDate.Value.Day; i <= endDay; i++)
                {
                    var day = new DateTime(DateTime.Now.Year, month, i);

                    if (day.DayOfWeek != DayOfWeek.Sunday)
                    {
                        monthWorkingDays.Add(i);
                    }
                }
            }

            return monthWorkingDays;
        }


        public int CalculateStartDay(IList<int> workingDays, IList<int> monthWorkingDays)
        {
            int startDay = 1;
            int endDay = workingDays.Count - 1;

            for (int i = 0; i <= endDay; i++)
            {
                if (workingDays[i] != monthWorkingDays[0])
                {
                    startDay++;
                }
                else
                {
                    break;
                }
            }

            return startDay;
        }

        public bool IsMonthInRange(StudentDataRequest request, int month)
        {
            int? startMonth = request.StartDate?.Month;
            int? endMonth = request.EndDate?.Month;

            if (startMonth == endMonth && startMonth == month)
            {
                return !(request.StartDate?.Day == 1 && request.EndDate?.Day == DateTime.DaysInMonth(request.EndDate?.Year ?? DateTime.Now.Year, month));
            }
            else if (startMonth == month)
            {
                return !(request.StartDate?.Day == 1);
            }
            else if (endMonth == month)
            {
                return !(request.EndDate?.Day == DateTime.DaysInMonth(request.EndDate?.Year ?? DateTime.Now.Year, month));
            }

            return false;
        }



        public List<IList<object>> FilterValues(List<IList<object>> combinedValues, StudentDataRequest request)
        {
            return combinedValues
                .Where((row, index) =>
                    (index == 0) || // Include header row
                    (
                        // Your filtering conditions go here based on request parameters and row values
                        (string.IsNullOrEmpty(request.Class) || string.Equals(row[0]?.ToString(), request.Class, StringComparison.OrdinalIgnoreCase)) &&
                        (string.IsNullOrEmpty(request.Name) || AreStringsEqual(row[1]?.ToString() ?? "", request.Name)) &&
                        (string.IsNullOrEmpty(request.Surname) || AreStringsEqual(row[2]?.ToString() ?? "", request.Surname)) &&
                        (string.IsNullOrEmpty(request.Code) || string.Equals(row[3]?.ToString(), request.Code, StringComparison.OrdinalIgnoreCase)) &&
                        (
                            (request.HasContract == null) ||
                            (request.HasContract == true && string.Equals(row[4]?.ToString(), "+", StringComparison.OrdinalIgnoreCase)) ||
                            (request.HasContract == false && string.IsNullOrEmpty(row[4]?.ToString()))
                        ) &&
                        (
                            (request.IsPaid == null) ||
                            (request.IsPaid == true && string.Equals(row[5]?.ToString(), "+", StringComparison.OrdinalIgnoreCase)) ||
                            (request.IsPaid == false && string.IsNullOrEmpty(row[5]?.ToString()))
                        )
                    )
                )
                .ToList();
        }

        public int Compare(string? x, string? y)
        {
            return x?.Length == y?.Length
                ? string.Compare(x, y, StringComparison.OrdinalIgnoreCase)
                : x?.Length.CompareTo(y?.Length) ?? 0;
        }

    }
}

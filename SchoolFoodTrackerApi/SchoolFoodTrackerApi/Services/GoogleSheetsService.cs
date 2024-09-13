using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using Newtonsoft.Json;
using SchoolFoodTrackerApi.Models;
using static SchoolFoodTrackerApi.Models.PersonSearchModel;

namespace SchoolFoodTrackerApi.Services
{
    public class GoogleSheetsService : IGoogleSheetsService
    {
        private readonly SheetsService _sheetsService;
        private readonly ISchoolFoodService _schoolFoodService;
        private readonly string _spreadsheetId;
        private readonly ILogger<GoogleSheetsService> _logger;
        private readonly Dictionary<int, string> _monthNames;

        public GoogleSheetsService(ILogger<GoogleSheetsService> logger, ISchoolFoodService schoolFoodService)
        {
            _schoolFoodService = schoolFoodService;
            _sheetsService = InitializeGoogleSheetsService();
            _spreadsheetId = "1FyoOOiT8x7VMFP4uXSgsfANtGA_gEH_lIDU90Q5-mpQ";
            _logger = logger;
            _monthNames = new Dictionary<int, string>
            {
                { 1, "January" },
                { 2, "February" },
                { 3, "March" },
                { 4, "April" },
                { 5, "May" },
                { 9, "September" },
                { 10, "October" },
                { 11, "November" },
                { 12, "December" }
            };
            _schoolFoodService = schoolFoodService;
        }
        public SheetsService InitializeGoogleSheetsService()
        {
            var clientSecretPath = "client_secret_621939025240-u45shg2lpm2e037gl4q6oft4q83qqfn5.apps.googleusercontent.com.json";

            ClientSecrets secrets;
            using (var stream = new FileStream(clientSecretPath, FileMode.Open, FileAccess.Read))
            {
                secrets = GoogleClientSecrets.FromStream(stream).Secrets;
            }

            var sheetsService = new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = GetCredential(secrets),
                ApplicationName = "SchoolFood",
            });

            return sheetsService;
        }

        private static UserCredential GetCredential(ClientSecrets secrets)
        {
            var credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                secrets,
                new[] { SheetsService.Scope.Spreadsheets },
                "user",
                CancellationToken.None,
                new FileDataStore("SchoolFoodTracker")
            ).Result;

            return credential;
        }

        private void ClearSheetExceptFirstRow(string sheetName)
        {
            // Определите диапазон для очистки (все строки, начиная со второй)
            var rangeToClear = $"{sheetName}!A2:Z";

            var clearRequest = _sheetsService.Spreadsheets.Values.Clear(new ClearValuesRequest(), _spreadsheetId, rangeToClear);
            clearRequest.Execute();

            _logger.LogInformation("Sheet {SheetName} cleared except for the first row.", sheetName);
        }

        private void UpdateDataMonth(int currentMonth)
        {
            var sheetName = _monthNames[currentMonth];
            var range = $"{sheetName}!A2";

            // Очищаем таблицу, кроме первой строки
            ClearSheetExceptFirstRow(sheetName);

            // Чтение данных из JSON файла
            var jsonData = System.IO.File.ReadAllText("schoolData.json");
            var studentDataList = JsonConvert.DeserializeObject<List<StudentData>>(jsonData);

            // Создание списка для хранения данных из Excel файлов
            var excelDataList = new List<StudentData>();

            // Чтение данных из каждого Excel файла
            foreach (var excelFilePath in _schoolFoodService.GetExcelFilePaths())
            {
                var excelStudentDataList = _schoolFoodService.ReadCsvFile(excelFilePath);
                excelDataList.AddRange(excelStudentDataList);
            }

            // Объединение данных из JSON и Excel
            var combinedDataList = studentDataList!
                .Concat(excelDataList!)
                .OrderBy(student => student.Klase, Comparer<string>.Create((x, y) => _schoolFoodService.Compare(x, y)))
                .ToList();

            var values = new List<IList<object>>();

            foreach (var studentData in combinedDataList)
            {
                if (!_schoolFoodService.IsExcludedClass(studentData.Klase))
                {
                    var rowValues = new List<object>
                    {
                        studentData.Klase,
                        studentData.Vards,
                        studentData.Uzvards,
                        studentData.Kods,
                        studentData.Ligums ?? "",
                        studentData.IrSamaksats ?? ""
                    };

                    values.Add(rowValues);
                }
            }

            var valueRange = new ValueRange
            {
                Values = values
            };

            var updateRequest = _sheetsService.Spreadsheets.Values.Update(valueRange, _spreadsheetId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            var updateResponse = updateRequest.Execute();

            _logger.LogInformation("Data successfully updated in the {SheetName} sheet.", sheetName);
        }

        public SearchResult FindAndPrintRow(string sheetName, int code)
        {

            SearchResult result = new SearchResult();

            int daysInMonth = DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month);

            var workingDays = new List<int>();

            for (int i = 1; i <= daysInMonth; i++)
            {
                var day = new DateTime(DateTime.Now.Year, DateTime.Now.Month, i);

                if (day.DayOfWeek != DayOfWeek.Sunday)
                {
                    workingDays.Add(i);
                }
            }

            var lastWorkingDay = workingDays.Count + 5;
            var range = $"{sheetName}!A1:{_schoolFoodService.ConvertIndexToColumnName(lastWorkingDay)}";

            var request = _sheetsService.Spreadsheets.Values.Get(_spreadsheetId, range);
            var response = request.Execute();
            var values = response.Values;

            if (values != null && values.Count > 0)
            {
                int maxColumnCount = values.Max(row => row.Count);

                foreach (var row in values)
                {
                    while (row.Count < maxColumnCount)
                    {
                        row.Add(null);
                    }
                }

                foreach (var row in values)
                {
                    if (row.Count > 3 && row[3] != null && !string.IsNullOrEmpty(row[3].ToString()) && row[3].ToString() == code.ToString())
                    {
                        _logger.LogInformation("Row data: {RowData}", string.Join(", ", row));

                        result = new SearchResult
                        {
                            PersonInfo = new PersonInfo
                            {
                                Class = row[0]?.ToString(),
                                Name = row[1]?.ToString(),
                                Surname = row[2]?.ToString(),
                                HasContract = row[4]?.ToString() == "+",
                                IsPaid = row[5]?.ToString() == "+"
                            }
                        };

                        var currentDayIndex = DateTime.Now.Day;

                        if (currentDayIndex <= row.Count)
                        {
                            int columnIndex = -1;

                            for (int i = 0; i < values[0].Count; i++)
                            {
                                if (values[0][i] != null && values[0][i].ToString() == currentDayIndex.ToString())
                                {
                                    columnIndex = i;
                                    break;
                                }
                            }

                            if (columnIndex != -1)
                            {
                                result.Status = "viss labi";

                                if (row[columnIndex] != null && row[columnIndex].ToString() == "+")
                                {
                                    result.Status = "Pirms tas atzimets";
                                    return result;
                                }

                                if (row[4] == null && row[5] == null)
                                {
                                    result.Status = "nav liguma un nav samaksats";
                                    return result;
                                }
                                else if (row[4] == null)
                                {
                                    result.Status = "nav liguma";
                                    return result;
                                }
                                else if (row[5] == null)
                                {
                                    result.Status = "nav samaksats";
                                    return result;
                                }

                                row[columnIndex] = "+";
                                var updateRange = $"{sheetName}!{_schoolFoodService.ConvertIndexToColumnName(columnIndex)}{values.IndexOf(row) + 1}";
                                var updateValues = new List<IList<object>> { new List<object> { "+" } };
                                var updateRequest = _sheetsService.Spreadsheets.Values.Update(new ValueRange { Values = updateValues }, _spreadsheetId, updateRange);
                                updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                                var updateResponse = updateRequest.Execute();

                                _logger.LogInformation("Successfully updated for day {CurrentDayIndex}.", currentDayIndex);
                                return result;
                            }
                            else
                            {
                                _logger.LogError("Error: Skipped day for {ColumnIndex} (weekend).", columnIndex);
                                result.Status = "Šodien nevar atzīmēt(sestdiena vai svētdiena)";
                                return result;
                            }
                        }
                        else
                        {
                            _logger.LogError("Error: Current day index exceeds the array size.");
                            result.Status = "Šadu dienu nav menesī";
                            return result;
                        }
                    }
                }
                _logger.LogInformation("Row not found.");
                result.Status = "Nav tadu cilvēku";
                return result;
            }
            else
            {
                _logger.LogInformation("Data not found.");
                result.Status = "Nav datu";
                return result;
            }
        }

        public List<IList<object>> RetrieveStudentDataFromSheets(StudentDataRequest request, int month)
        {
            int daysInMonth = DateTime.DaysInMonth(DateTime.Now.Year, month);
            var workingDays = _schoolFoodService.CalculateWorkingDaysInSpecifiedRange(request, month, daysInMonth);

            if (!_schoolFoodService.IsMonthInRange(request, month))
            {
                return RetrieveStudentDataForEntireMonth(request, month, workingDays);
            }
            else
            {
                return RetrieveStudentDataForSpecifiedRange(request, month, workingDays, daysInMonth);
            }
        }

        private List<IList<object>> RetrieveStudentDataForEntireMonth(StudentDataRequest request, int month, IList<int> workingDays)
        {
            var lastWorkingDay = workingDays.Count + 5; // Adjusted to consider weekends
            var sheetName = _monthNames[month];
            var range = $"{sheetName}!A1:{_schoolFoodService.ConvertIndexToColumnName(lastWorkingDay)}";

            return RetrieveDataFromSheet(range, sheetName, request);
        }

        private List<IList<object>> RetrieveStudentDataForSpecifiedRange(StudentDataRequest request, int month, IList<int> workingDays, int daysInMonth)
        {
            var monthWorkingDays = _schoolFoodService.CalculateWorkingDaysInSpecifiedRange(request, month, daysInMonth);
            var startDay = _schoolFoodService.CalculateStartDay(workingDays, monthWorkingDays);
            var sheetName = _monthNames[month];

            SpreadsheetsResource.ValuesResource.BatchGetRequest batchRequest = _sheetsService.Spreadsheets.Values.BatchGet(_spreadsheetId);
            batchRequest.Ranges = new List<string>
            {
                $"{sheetName}!A:F",
                $"{sheetName}!{_schoolFoodService.ConvertIndexToColumnName(startDay + 5)}:{_schoolFoodService.ConvertIndexToColumnName(monthWorkingDays.Count + startDay + 4)}"
            };

            try
            {
                BatchGetValuesResponse batchResponse = batchRequest.Execute();
                IList<ValueRange> valueRanges = batchResponse.ValueRanges;

                List<IList<object>> combinedValues = [];

                foreach (var valueRange in valueRanges)
                {
                    IList<IList<object>> values = valueRange.Values ?? new List<IList<object>>();

                    if (combinedValues.Count == 0)
                    {
                        combinedValues.AddRange(values);
                    }
                    else
                    {
                        for (int i = 0; i < values.Count; i++)
                        {
                            if (combinedValues.Count <= i)
                            {
                                combinedValues.Add(new List<object>(values[i]));
                            }
                            else
                            {
                                combinedValues[i] = combinedValues[i].Concat(values[i]).ToList();
                            }
                        }
                    }
                }

                int maxColumnCount = combinedValues.Count != 0 ? combinedValues.Max(row => row.Count) : 0;

                foreach (var row in combinedValues)
                {
                    while (row.Count < maxColumnCount)
                    {
                        row.Add("");
                    }
                }

                combinedValues = _schoolFoodService.FilterValues(combinedValues, request);

                return combinedValues;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error retrieving data from {sheetName} in Google Sheets: {ex.Message}");
                return [];
            }
        }


        private List<IList<object>> RetrieveDataFromSheet(string range, string sheetName, StudentDataRequest request)
        {
            SpreadsheetsResource.ValuesResource.GetRequest getRequest = _sheetsService.Spreadsheets.Values.Get(_spreadsheetId, range);
            ValueRange response;

            try
            {
                response = getRequest.Execute();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error retrieving student data from {SheetName} in Google Sheets.", sheetName);

                return [];
            }

            List<IList<object>> values = response.Values?.ToList() ?? new List<IList<object>>();

            // Example data processing and filtering logic (replace with your own logic)
            values = _schoolFoodService.FilterValues(values, request);


            return values;
        }

        public List<string> GetAllClasses()
        {
            try
            {
                var allClasses = new List<string>();

                foreach (var month in _monthNames.Values)
                {
                    var range = $"{month}!A2:A";
                    var request = _sheetsService.Spreadsheets.Values.Get(_spreadsheetId, range);
                    var response = request.Execute();
                    var values = response.Values;

                    if (values != null && values.Count > 0)
                    {
                        // Регистронезависимое преобразование строк к верхнему регистру
                        var classes = values.Select(row => row[0]?.ToString()?.ToUpper()).ToList();
                        allClasses.AddRange(classes.Cast<string>());
                    }
                }

                if (allClasses.Count > 0)
                {
                    // Удалите дубликаты, если необходимо (с учетом регистра)
                    var uniqueClasses = allClasses.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                    return uniqueClasses;
                }
                else
                {
                    return [];
                }
            }
            catch (Exception ex)
            {
                // Вместо BadRequest, можете возвращать пустой список или другой результат, 
                // в зависимости от вашей логики обработки ошибок.
                throw new Exception($"An error occurred: {ex.Message}");
            }
        }
        public void UpdateData(DateTime startDate, DateTime endDate)
        {
            try
            {
                System.IO.File.WriteAllText("schoolData.json", "[]");

                for (DateTime currentDate = startDate; currentDate <= endDate; currentDate = currentDate.AddMonths(1))
                {
                    int currentMonth = currentDate.Month;

                    // Пропустить обработку для месяцев 6, 7 и 8 при startMonth = 1 и endMonth = 12
                    if (currentMonth >= 6 && currentMonth <= 8)
                    {
                        continue;
                    }

                    UpdateDataMonth(currentMonth);
                }
            }
            catch (Exception ex)
            {
                // Вместо прокидывания исключения можно выполнить логирование ошибки или другую логику обработки ошибок
                throw new Exception($"An error occurred during data update: {ex.Message}");
            }
        }
        public void UpdateHeaders(int startMonth, int endMonth)
        {
            try
            {
                for (int month = startMonth; month <= endMonth; month++)
                {
                    if (month < 6 || month > 8)
                    {
                        var sheetName = _monthNames[month];
                        var range = $"{sheetName}!A1";

                        var values = new List<IList<object>>();

                        var headerValues = new List<object>
                        {
                            "Klase", "Vārds", "Uzvārds", "Kods", "Līgums", "Ir Samaksāts"
                        };

                        for (int i = 1; i <= DateTime.DaysInMonth(DateTime.Now.Year, month); i++)
                        {
                            var day = new DateTime(DateTime.Now.Year, month, i);

                            if (day.DayOfWeek != DayOfWeek.Sunday)
                            {
                                headerValues.Add(i.ToString());
                            }
                        }
                        values.Add(headerValues);

                        var valueRange = new ValueRange
                        {
                            Values = values
                        };

                        var updateRequest = _sheetsService.Spreadsheets.Values.Update(valueRange, _spreadsheetId, range);
                        updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                        updateRequest.Execute();

                        _logger.LogInformation("Data successfully updated in the {SheetName} sheet.", sheetName);
                    }
                }
            }
            catch (Exception ex)
            {
                // Вместо прокидывания исключения можно выполнить логирование ошибки или другую логику обработки ошибок
                throw new Exception($"An error occurred during updating headers: {ex.Message}");
            }
        }
    }
}

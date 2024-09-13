using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using OfficeOpenXml;
using System.ComponentModel.DataAnnotations;
using System.Globalization;
using CsvHelper;

using SchoolFoodTrackerApi.Models;
using SchoolFoodTrackerApi.Services;
using static SchoolFoodTrackerApi.Models.PersonSearchModel;

namespace SchoolFoodTrackerApi.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class SchoolFoodController : ControllerBase
    {
        private readonly Dictionary<int, string> _monthNames;
        private readonly ISchoolFoodService _schoolFoodService;
        private readonly IGoogleSheetsService _googleSheetsService;

        public SchoolFoodController(ISchoolFoodService schoolFoodService, IGoogleSheetsService googleSheetsService)
        {
            _schoolFoodService = schoolFoodService;
            _googleSheetsService = googleSheetsService;  // Обратите внимание, что теперь это IGoogleSheetsService

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
        }


        [HttpGet("get-all-classes")]
        public IActionResult GetAllClasses()
        {
            try
            {
                var uniqueClasses = _googleSheetsService.GetAllClasses();

                if (uniqueClasses.Count > 0)
                {
                    return Ok(uniqueClasses);
                }
                else
                {
                    return NotFound("No classes found in the spreadsheet.");
                }
            }
            catch (Exception ex)
            {
                return BadRequest(new { success = false, message = ex.Message });
            }
        }



        [HttpGet("update-data")]
        public IActionResult UpdateData(DateTime? startDate, DateTime? endDate)
        {
            try
            {
                if (!startDate.HasValue)
                {
                    startDate = DateTime.Now.Date; // Используйте текущую дату с началом дня
                }

                if (!endDate.HasValue)
                {
                    endDate = startDate.Value; // Если конечная дата не указана, используйте начальную
                }

                if (startDate > endDate)
                {
                    return BadRequest(new { success = false, message = "End date should be greater than or equal to start date." });
                }

                _googleSheetsService.UpdateData(startDate.Value, endDate.Value);

                return Ok(new { success = true, message = "Data successfully updated." });
            }
            catch (Exception ex)
            {
                return BadRequest(new { success = false, message = ex.Message });
            }
        }


        [HttpGet("find/{code}")]
        public IActionResult FindAndPrintRow(int code)
        {
            try
            {
                int currentMonth = DateTime.Now.Month;
                SearchResult finded = _googleSheetsService.FindAndPrintRow(_monthNames[currentMonth], code);
                return Ok(finded); 
            }
            catch (Exception ex)
            {
                SearchResult error = new SearchResult();
                error.Status = $"An error occurred: {ex.Message}";
                return BadRequest(error);
            }
        }

        [HttpGet("update-headers")]
        public IActionResult UpdateHeaders(int? startMonth, int? endMonth)
        {
            try
            {
                if (!startMonth.HasValue)
                {
                    startMonth = 1;
                }

                if (!endMonth.HasValue)
                {
                    endMonth = 12;
                }

                if (startMonth < 1 || startMonth > 12 || endMonth < 1 || endMonth > 12 || startMonth > endMonth)
                {
                    return BadRequest("Invalid month range.");
                }

                _googleSheetsService.UpdateHeaders(startMonth.Value, endMonth.Value);

                return Ok("Headers successfully updated.");
            }
            catch (Exception ex)
            {
                return BadRequest($"An error occurred: {ex.Message}");
            }
        }

        [HttpPost("get-student-data")]
        public IActionResult GetStudentData([FromBody] StudentDataRequest request)
        {
            // Проверка наличия параметров запроса и вызов метода для получения данных из Google Sheets
            if (request == null)
            {
                return BadRequest("Request body is empty.");
            }

            DateTime startDate = request.StartDate ?? new DateTime(DateTime.Now.Year, 1, 1);
            DateTime endDate = request.EndDate ?? new DateTime(DateTime.Now.Year, 12, 1);

            var result = new List<object>();

                for (DateTime date = startDate; date <= endDate; date = date.AddMonths(1))
                {
                    // Пропускать летние месяцы (июнь, июль, август)
                    if (date.Month == 6 || date.Month == 7 || date.Month == 8)
                        continue;

                    var monthSheetName = date.ToString("MMMM");
                    var studentData = _googleSheetsService.RetrieveStudentDataFromSheets(request, date.Month);

                    if (studentData == null)
                    {
                        return StatusCode(500, $"Failed to retrieve student data from {monthSheetName}.");
                    }

                    result.Add(new
                    {
                        Month = monthSheetName,
                        Data = studentData
                    });
                }

            byte[] excelData = _schoolFoodService.CreateExcelFile(result);

            // Отправка Excel-файла в ответе на запрос
            return File(excelData, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "StudentData.xlsx");
        }      
    }
}

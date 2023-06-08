using ExportQueryToExcelStream;
using Microsoft.AspNetCore.Mvc;
using System.Data;

namespace ExportListToExcel.WebApi.Controllers {

    [ApiController]
    [Route("[controller]")]
    public class WeatherForecastController : ControllerBase {

        private static readonly string[] Summaries = new[]
        {
            "Freezing", "Bracing", "Chilly", "Cool", "Mild", "Warm", "Balmy", "Hot", "Sweltering", "Scorching"
        };

        private readonly ILogger<WeatherForecastController> _logger;

        public WeatherForecastController(ILogger<WeatherForecastController> logger)
        {
            _logger = logger;
        }

        [HttpGet(Name = "GetWeatherForecast")]
        public IEnumerable<WeatherForecast> Get()
        {
            return Enumerable.Range(1, 5).Select(index => new WeatherForecast
            {
                Date = DateOnly.FromDateTime(DateTime.Now.AddDays(index)),
                TemperatureC = Random.Shared.Next(-20, 55),
                Summary = Summaries[Random.Shared.Next(Summaries.Length)]
            })
            .ToArray();
        }

        [HttpGet]
        [Route("DownloadExcel")]
        public async Task<IActionResult> DownloadExcel()
        {
            StreamExcelConverter converter = new(Get().AsQueryable());
            var result = new FileStreamResult(converter.ToExcelStream(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            {
                FileDownloadName = "MyFileName.xlsx"
            };

            return result;
        }
    }
}
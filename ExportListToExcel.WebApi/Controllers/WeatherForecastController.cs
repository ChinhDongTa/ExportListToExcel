using ClosedXML.Excel;
using ExportQueryToExcelStream;
using Microsoft.AspNetCore.Mvc;
using System.Data;
using System.IO;

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
            //IQueryable queryable = ;
            StreamExcelConverter converter = new(Get().AsQueryable());
            var result = new FileStreamResult(converter.ToExcelStream(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            {
                FileDownloadName = "MyFileName.xlsx"
            };

            return result;
        }

        //#region Private Methods

        //private static FileStreamResult ExportExcel(IQueryable queryable)
        //{
        //    var talbe = Convert2Datatable(queryable.AsQueryable());
        //    using XLWorkbook wb = new();
        //    wb.Worksheets.Add(talbe);
        //    //Không cần sử dụng "using(MemoryStream stream = new())"
        //    MemoryStream stream = new();
        //    wb.SaveAs(stream);
        //    stream.Position = 0;
        //    var result = new FileStreamResult(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        //    {
        //        FileDownloadName = "MyFileName.xlsx"
        //    };

        //    return result;
        //}

        //private static DataTable Convert2Datatable(IQueryable query)
        //{
        //    if (query != null)
        //    {
        //        DataTable dt = new(query.ElementType.Name);
        //        var columns = GetProperties(query.ElementType);
        //        dt.Columns.AddRange(ToDataColumn(columns));
        //        foreach (var item in query)
        //        {
        //            var row = new List<string>();
        //            //DataRow row2 = dt.NewRow();
        //            foreach (var column in columns)
        //            {
        //                var value = GetValue(item, column);
        //                row.Add(value != null ? value.ToString() : "null");
        //            }
        //            dt.Rows.Add(row.ToArray());
        //        }
        //        return dt;
        //    }
        //    return new DataTable("NotFound");
        //}

        //private static DataColumn[] ToDataColumn(IEnumerable<string> columns)
        //{
        //    var dataCol = new List<DataColumn>();
        //    foreach (var column in columns)
        //        dataCol.Add(new DataColumn(column));
        //    return dataCol.ToArray();
        //}

        //private static object GetValue(object item, string column)
        //{
        //    return item.GetType()?.GetProperty(column)?.GetValue(item) ?? "Null";
        //}

        ///// <summary>
        ///// Get fileds name
        ///// </summary>
        ///// <param name="elementType"></param>
        ///// <returns></returns>
        //private static IEnumerable<string> GetProperties(Type elementType)
        //{
        //    return elementType.GetProperties().Select(p => p.Name);
        //}

        //#endregion Private Methods
    }
}
using ClosedXML.Excel;
using System.Data;

namespace ExportQueryToExcelStream {

    public class StreamExcelConverter {
        private readonly IQueryable? _queryable;

        public StreamExcelConverter(IQueryable queryable)
        {
            _queryable = queryable;
        }

        public Stream ToExcelStream()
        {
            var talbe = Convert2Datatable(_queryable);
            using XLWorkbook wb = new();
            wb.Worksheets.Add(talbe);
            //Không cần sử dụng "using(MemoryStream stream = new())"
            MemoryStream stream = new();
            wb.SaveAs(stream);
            stream.Position = 0;

            return stream;
        }

        #region Private Methods

        private static DataTable Convert2Datatable(IQueryable query)
        {
            if (query != null)
            {
                DataTable dt = new(query.ElementType.Name);
                var columns = GetProperties(query.ElementType);
                dt.Columns.AddRange(ToDataColumn(columns));
                foreach (var item in query)
                {
                    var row = new List<string>();
                    //DataRow row2 = dt.NewRow();
                    foreach (var column in columns)
                    {
                        var value = GetValue(item, column);
                        row.Add(value != null ? value.ToString() : "null");
                    }
                    dt.Rows.Add(row.ToArray());
                }
                return dt;
            }
            return new DataTable("NotFound");
        }

        private static DataColumn[] ToDataColumn(IEnumerable<string> columns)
        {
            var dataCol = new List<DataColumn>();
            foreach (var column in columns)
                dataCol.Add(new DataColumn(column));
            return dataCol.ToArray();
        }

        private static object GetValue(object item, string column)
        {
            return item.GetType()?.GetProperty(column)?.GetValue(item) ?? "Null";
        }

        /// <summary>
        /// Get fileds name
        /// </summary>
        /// <param name="elementType"></param>
        /// <returns></returns>
        private static IEnumerable<string> GetProperties(Type elementType)
        {
            return elementType.GetProperties().Select(p => p.Name);
        }

        #endregion Private Methods
    }
}
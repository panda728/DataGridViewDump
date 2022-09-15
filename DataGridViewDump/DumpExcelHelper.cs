using System.Globalization;
using FakeExcelSerializer;

namespace DataGridViewDump
{
    public static class DumpExcelHelper
    {
        readonly static IExcelSerializerProvider _dataGridViewExcelProvider
            = ExcelSerializerProvider.Create(
                new[] { new DataGridViewExcelSerializer() },
                new[] { ExcelSerializerProvider.Default });

        internal sealed class DataGridViewExcelSerializer : IExcelSerializer<DataGridViewRow>
        {
            public void WriteTitle(ref ExcelSerializerWriter writer, DataGridViewRow value, ExcelSerializerOptions options, string name)
            {
                var serializer = options.GetRequiredSerializer<string>();
                var headerTexts = value.Cells.OfType<DataGridViewCell>()
                    .OrderBy(c => c.OwningColumn.DisplayIndex)
                    .Select(c => c.OwningColumn.HeaderText);
                foreach (var text in headerTexts)
                {
                    serializer.Serialize(ref writer, text, options);
                }
            }

            public void Serialize(ref ExcelSerializerWriter writer, DataGridViewRow value, ExcelSerializerOptions options)
            {
                var serializer = options.GetRequiredSerializer<object>();
                var cells = value.Cells.OfType<DataGridViewCell>()
                    .OrderBy(c => c.OwningColumn.DisplayIndex);
                foreach (var cell in cells)
                {
                    serializer.Serialize(ref writer, cell.Value, options);
                }
            }
        }

        public static void ToExcelFile<T>(
            this IEnumerable<T> value,
            string fileName,
            bool hasHeaderRecord = true,
            string[]? headerTitles = null,
            bool autoFitColumns = true
        )
        {
            if (value == null || !value.Any())
                return;

            var newConfig = ExcelSerializerOptions.Default with
            {
                Provider = _dataGridViewExcelProvider,
                CultureInfo = CultureInfo.CurrentCulture,
                HasHeaderRecord = hasHeaderRecord || (headerTitles?.Any() ?? false),
                HeaderTitles = headerTitles ?? Array.Empty<string>(),
                AutoFitColumns = autoFitColumns,
                NumberFormat = "#,##0.000;[Red]\\-#,##0.000",
            };
            ExcelSerializer.ToFile(value, fileName, newConfig);
        }

        public static void ToExcelFile(this DataGridView dataGridView, string fileName)
        {
            var rows = dataGridView.Rows.OfType<DataGridViewRow>();
            if (rows == null || !rows.Any())
                return;

            var headerTitles = dataGridView.Columns.OfType<DataGridViewColumn>()
                .OrderBy(c => c.DisplayIndex)
                .Select(c => c.HeaderText)
                .ToArray();

            ToExcelFile(rows, fileName, true, headerTitles, true);
        }
    }
}

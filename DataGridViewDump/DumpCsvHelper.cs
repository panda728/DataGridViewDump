using System.Globalization;
using System.Text;
using FakeCsvSerializer;

namespace DataGridViewDump
{
    public static class DumpCsvHelper
    {
        readonly static ICsvSerializerProvider _dataGridViewCsvProvider
            = CsvSerializerProvider.Create(
                new[] { new DataGridViewCsvSerializer() },
                new[] { CsvSerializerProvider.Default });

        internal sealed class DataGridViewCsvSerializer : ICsvSerializer<DataGridViewRow>
        {
            public void WriteTitle(ref CsvSerializerWriter writer, DataGridViewRow value, CsvSerializerOptions options, string name)
            {
                var serializer = options.GetRequiredSerializer<string>();
                var headerTexts = value.Cells.OfType<DataGridViewCell>()
                    .OrderBy(c => c.OwningColumn.DisplayIndex)
                    .Select(c => c.OwningColumn.HeaderText);
                foreach (var text in headerTexts)
                {
                    writer.WriteDelimiter();
                    serializer.Serialize(ref writer, text, options);
                }
            }
            public void Serialize(ref CsvSerializerWriter writer, DataGridViewRow value, CsvSerializerOptions options)
            {
                var serializer = options.GetRequiredSerializer<object>();
                var cells = value.Cells.OfType<DataGridViewCell>()
                    .OrderBy(c => c.OwningColumn.DisplayIndex);
                foreach (var cell in cells)
                {
                    writer.WriteDelimiter();
                    serializer.Serialize(ref writer, cell.Value, options);
                }
            }
        }

        public static void ToCsvFile<T>(
            this IEnumerable<T> value,
            string fileName,
            bool hasHeaderRecord = true,
            string[]? headerTitles = null,
            Encoding? encoding = null
        )
        {
            if (value == null || !value.Any())
                return;

            var newConfig = CsvSerializerOptions.Default with
            {
                Provider = _dataGridViewCsvProvider,
                CultureInfo = CultureInfo.CurrentCulture,
                Encoding = encoding ?? Encoding.UTF8,
                HasHeaderRecord = hasHeaderRecord || (headerTitles?.Any() ?? false),
                HeaderTitles = headerTitles ?? Array.Empty<string>(),
            };
            CsvSerializer.ToFile(value, fileName, newConfig);
        }

        public static void ToCsvFile(this DataGridView dataGridView, string fileName, Encoding? encoding = null)
        {
            var rows = dataGridView.Rows.OfType<DataGridViewRow>();
            if (rows == null || !rows.Any())
                return;

            var headerTitles = dataGridView.Columns.OfType<DataGridViewColumn>()
                .OrderBy(c => c.DisplayIndex)
                .Select(c => c.HeaderText)
                .ToArray();

            ToCsvFile(rows, fileName, true, headerTitles, encoding);
        }
    }
}
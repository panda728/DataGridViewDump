using FakeExcelSerializer;
using System.Diagnostics;
using System.Globalization;

namespace DataGridViewDump
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            for (int i = 0; i < 10; i++)
            {
                var row = dataSet11.Table1.NewTable1Row();
                row.Col1 = $"Col1-{i}";
                row.Col2 = $"Col2-{i}";
                row.Col3 = i;
                dataSet11.Table1.AddTable1Row(row);
            }
            dataSet11.AcceptChanges();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                saveFileDialog1.FileName = "datagridviewdump.xlsx";
                if (saveFileDialog1.ShowDialog() != DialogResult.OK)
                    return;

                ExportDataGridViewToExcel(dataGridView1, saveFileDialog1.FileName);

                var psi = new ProcessStartInfo();
                psi.UseShellExecute = true;
                psi.FileName = saveFileDialog1.FileName;

                using (var p = Process.Start(psi))
                    p.WaitForExit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        readonly IExcelSerializerProvider _dataGridViewExcelProvider
            = ExcelSerializerProvider.Create(
                new[] { new DataGridViewExcelSerializer() },
                new[] { ExcelSerializerProvider.Default });

        private void ExportDataGridViewToExcel(DataGridView dataGridView, string fileName)
        {
            var titles = dataGridView.Columns.Cast<DataGridViewColumn>()
                .Select(c => c.HeaderText).ToArray();

            var newConfig = ExcelSerializerOptions.Default with
            {
                CultureInfo = CultureInfo.CurrentCulture,
                Provider = _dataGridViewExcelProvider,
                HasHeaderRecord = true,
                HeaderTitles = titles,
                AutoFitColumns = false,
            };

            ExcelSerializer.ToFile(dataGridView1.Rows.Cast<DataGridViewRow>(), fileName, newConfig);
        }
    }

    public class DataGridViewExcelSerializer : IExcelSerializer<DataGridViewRow>
    {
        public void Serialize(ref ExcelSerializerWriter writer, DataGridViewRow value, ExcelSerializerOptions options)
        {
            var serializer = options.GetRequiredSerializer<object>();
            var columns = value.Cells.Cast<DataGridViewCell>().ToArray().AsSpan();
            foreach (var c in columns)
                serializer.Serialize(ref writer, c.Value, options);
        }
    }
}
using System.Diagnostics;

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

        void button1_Click(object sender, EventArgs e)
        {
            try
            {
                saveFileDialog1.FileName = "datagridviewdump.xlsx";
                if (saveFileDialog1.ShowDialog() != DialogResult.OK)
                    return;

                dataGridView1.ToExcelFile(saveFileDialog1.FileName);

                var psi = new ProcessStartInfo
                {
                    UseShellExecute = true,
                    FileName = saveFileDialog1.FileName
                };

                using (var p = Process.Start(psi))
                {
                    p.WaitForExit();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
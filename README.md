# DataGridViewDump

DataGridView to ExcelFile demo.

## Sample

![image](https://user-images.githubusercontent.com/16958552/186148469-1d721265-357f-4f95-8ba7-0ed059f24e76.png)

![image](https://user-images.githubusercontent.com/16958552/186148816-5c4ac74e-59df-46fd-91c6-7446cfa830a1.png)

## library
https://github.com/panda728/FakeExcelSerializer

## Excel file creation process

~~~
ExportDataGridViewToExcel(dataGridView1, saveFileDialog1.FileName);
~~~

## Overview

~~~
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
~~~

~~~
readonly IExcelSerializerProvider _dataGridViewExcelProvider
    = ExcelSerializerProvider.Create(
        new[] { new DataGridViewExcelSerializer() },
        new[] { ExcelSerializerProvider.Default });

~~~

~~~
public class DataGridViewExcelSerializer : IExcelSerializer<DataGridViewRow>
{
    public void WriteTitle(ref ExcelSerializerWriter writer, DataGridViewRow value, ExcelSerializerOptions options, string name)
    {
        var serializer = options.GetRequiredSerializer<object>();
        var columns = value.Cells.Cast<DataGridViewCell>().ToArray().AsSpan();
        foreach (var c in columns)
            serializer.Serialize(ref writer, c.OwningColumn.HeaderText, options);
    }
    public void Serialize(ref ExcelSerializerWriter writer, DataGridViewRow value, ExcelSerializerOptions options)
    {
        var serializer = options.GetRequiredSerializer<object>();
        var columns = value.Cells.Cast<DataGridViewCell>().ToArray().AsSpan();
        foreach (var c in columns)
            serializer.Serialize(ref writer, c.Value, options);
    }
}
~~~

## License
This library is licensed under the MIT License.

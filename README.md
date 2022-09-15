# DataGridViewDump

DataGridView to ExcelFile demo.

## Sample

![image](https://user-images.githubusercontent.com/16958552/186148469-1d721265-357f-4f95-8ba7-0ed059f24e76.png)

![image](https://user-images.githubusercontent.com/16958552/186148816-5c4ac74e-59df-46fd-91c6-7446cfa830a1.png)

## library
https://github.com/panda728/FakeExcelSerializer

## Excel file creation process

~~~
    saveFileDialog1.FileName = "datagridviewdump.xlsx";
    if (saveFileDialog1.ShowDialog() != DialogResult.OK)
        return;

    dataGridView1.ToExcelFile(saveFileDialog1.FileName);
~~~

## Overview

DumpExcelHelper.cs
~~~
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
~~~

## License
This library is licensed under the MIT License.

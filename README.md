# Excel.Helper
Repository for creating excel files from C# lists, and reading excel files to C# lists. This library uses ClosedXML to read from and write to excel sheets.

# Usage
Read from Excel File to a C# list
```
class Person 
{
  public int Id { get; set; }
  public string Name { get; set; }
}

List<Person> people = await ExcelReader.ReadExcelFile<Person>("EXCEL_FILE_NAME.xlsx");

// You can use people here
```

Write to a Excel File from a C# list
```
class Person 
{
  public int Id { get; set; }
  public string Name { get; set; }
}

byte[] excelBytes = await ExcelBuilder.BuildExcelFile(people);

// Returns a byte array, you can use that to
// write to a file or send over the network etc.

File.WriteAllBytesAsync("MY_EXCEL_FILE.xlsx", excelBytes);
```

Should note that this library only works with Office 2007+ format or *.xlsx* extensions because *ClosedXML* does not support *xls* extensions.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace Excel.Helper
{
    public static class ExcelReader
    {
        public static async Task<List<T>> ReadExcelFile<T>(byte[] fileContent, bool hasHeader = true)
        {
            await using var memoryStream = new MemoryStream(fileContent);

            return await ReadExcelFile<T>(memoryStream, hasHeader);
        }

        public static async Task<List<T>> ReadExcelFile<T>(string path, bool hasHeader = true)
        {
            await using var stream = File.OpenRead(path);

            return await ReadExcelFile<T>(stream, hasHeader);
        }

        public static async Task<List<T>> ReadExcelFile<T>(Stream stream, bool hasHeader = true)
        {
            using var workbook = new XLWorkbook(stream);
            var worksheet = workbook.Worksheets.First();
            var rows = hasHeader
                ? worksheet.Rows().Skip(1)
                : worksheet.Rows();
            
            var objList = new List<T>();
            var invalidColumns = new List<InvalidColumn>();
            
            if (typeof(T).IsPrimitive || typeof(T) == typeof(string))
            {
                foreach (var row in rows)
                {
                    try
                    {
                        var cell = row.Cell(1);
                        var obj = cell.GetValue<T>();
                        objList.Add(obj);
                    }
                    catch (FormatException e)
                    {
                        invalidColumns.Add(new InvalidColumn()
                        {
                            Column = 1,
                            Row = row.RowNumber(),
                            Exception = e,
                            Value = row.Cell(1).Value,
                            ConversionType = typeof(T)
                        });
                    }
                }
            }
            else
            {
                var properties = typeof(T)
                    .GetProperties();
                
                foreach (var row in rows)
                {
                    var obj = Activator.CreateInstance<T>();

                    for (var i = 0; i < properties.Length; i++)
                    {
                        try
                        {
                            var cell = row.Cell(i + 1);
                            var property = properties[i];
                            var castedValue = Convert.ChangeType(cell.Value, property.PropertyType);
                            property.SetValue(obj, castedValue);
                        }
                        catch (FormatException e)
                        {
                            invalidColumns.Add(new InvalidColumn()
                            {
                                Row = row.RowNumber(),
                                Column = i + 1,
                                Value = row.Cell(i + 1),
                                ConversionType = properties[i].PropertyType,
                                Exception = e
                            });
                        }
                    }

                    objList.Add(obj);
                }
            }

            if (invalidColumns.Any())
            {
                throw new InvalidExcelException(invalidColumns);
            }

            return objList;
        }
    }
}
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

            var properties = typeof(T)
                .GetProperties();

            var objList = new List<T>();
            var rows = hasHeader
                ? worksheet.Rows().Skip(1)
                : worksheet.Rows();

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
                    catch (FormatException)
                    {
                        throw new InvalidExcelException($"Cell value in Row : {row.RowNumber()}, Column : { 1 }, " +
                                                        $"Value : { row.Cell(1).Value } could not be converted to " +
                                                        $"the type : { typeof(T).Name }");
                    }
                }
            }
            else
            {
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
                        catch (FormatException)
                        {
                            throw new InvalidExcelException($"Cell value in Row : {row.RowNumber()}, Column : { i + 1 }, " +
                                                            $"Value : { row.Cell(i + 1).Value } could not be converted to " +
                                                            $"the type : { properties[i].PropertyType.Name }");
                        }
                    }

                    objList.Add(obj);
                }
            }

            return objList;
        }
    }
}
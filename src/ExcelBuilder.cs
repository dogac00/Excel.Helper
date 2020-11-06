using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using ClosedXML.Excel;
using Pluralize.NET.Core;

namespace Excel.Helper
{
    public static class ExcelBuilder
    {
        private static readonly Pluralizer _pluralizer = new Pluralizer();
        
        public static Task<byte[]> BuildExcelFile<T>(IEnumerable<T> list)
        {
            if (list == null)
                throw new ArgumentNullException("List parameter cannot be null.");
            
            var plural = _pluralizer.Pluralize(typeof(T).Name);
            
            return BuildExcelFile(list, plural);
        }
        
        public static async Task<byte[]> BuildExcelFile<T>(IEnumerable<T> list, string worksheetName)
        {
            if (list == null)
                throw new ArgumentNullException("List parameter cannot be null.");
            if (string.IsNullOrEmpty(worksheetName))
                throw new ArgumentException("You must enter worksheet name or use the default worksheet name.");
                
            using var workbook = new XLWorkbook();
            await using var stream = new MemoryStream();
            
            var dt = CreateDataTable(list);
            workbook.Worksheets.Add(dt, worksheetName);
            
            workbook.SaveAs(stream);
            return stream.ToArray();
        }

        private static DataTable CreateDataTable<T>(IEnumerable<T> list)
        {
            var dt = new DataTable();
            var properties = typeof(T)
                .GetProperties(BindingFlags.Public | BindingFlags.Instance);

            foreach (var property in properties)
            {
                dt.Columns.Add(property.Name, property.PropertyType);
                dt.Columns[property.Name].Caption = property.Name;
            }
            
            foreach (var element in list)
            {
                var row = dt.NewRow();

                foreach (var property in properties)
                {
                    row[property.Name] = property.GetValue(element);
                }

                dt.Rows.Add(row);
            }

            return dt;
        }
    }
}
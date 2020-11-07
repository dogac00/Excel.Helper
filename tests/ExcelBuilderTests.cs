using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Xunit;

namespace Excel.Helper.Tests
{
    public class ExcelBuilderTests
    {
        private string GetExcelsFolderPath(string fileName)
        {
            return $"../../../Excels/{fileName}";
        }
        
        [Fact]
        public async Task ShouldBuildExcel_WithPeopleList()
        {
            var people = new List<Person>
            {
                new Person {Id = 5, Name = "ABC"},
                new Person {Id = 77, Name = "DEF"},
                new Person {Id = 99, Name = "GHJ"}
            };

            var file = await ExcelBuilder.BuildExcelFile(people);
            var path = GetExcelsFolderPath("PeopleExcel.xlsx");
            
            await File.WriteAllBytesAsync(path, file);
        }
        
        [Fact]
        public async Task ShouldBuildExcel_WithStringList()
        {
            var textList = new List<string>
            {
                "TEXT 1",
                "TEXT 2",
                "TEXT 3"
            };

            var file = await ExcelBuilder.BuildExcelFile(textList);
            
            var path = GetExcelsFolderPath("StringsExcel.xlsx");

            await File.WriteAllBytesAsync(path, file);
        }
        
        [Fact]
        public async Task ShouldBuildExcel_WithDynamicList()
        {
            var textList = new List<dynamic>
            {
                new Person {Id = 12345, Name = "Dogac"},
                "Dogac2",
                34567
            };

            var file = await ExcelBuilder.BuildExcelFile(textList);
            
            var path = GetExcelsFolderPath("DynamicExcel.xlsx");

            await File.WriteAllBytesAsync(path, file);
        }
    }
}
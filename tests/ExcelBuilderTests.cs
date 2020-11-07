using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Xunit;

namespace Excel.Helper.Tests
{
    public class ExcelBuilderTests
    {
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
            
            await File.WriteAllBytesAsync("../../../Excels/SampleFile.xlsx", file);
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

            var textWritten = await ExcelReader.ReadExcelFile<string>(file);
            
            Assert.Equal(textList, textWritten);
        }
    }
}
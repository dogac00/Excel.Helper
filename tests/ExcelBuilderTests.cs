using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Excel.Helper.Tests.Types;
using Xunit;
using static Excel.Helper.Tests.ExcelFilesUtil;

namespace Excel.Helper.Tests
{
    public class ExcelBuilderTests
    {
        [Fact]
        public async Task BuildExcelFile_WhenValidPeopleListIsGiven_ShouldBuildExcelFile()
        {
            var people = new List<Person>
            {
                new Person {Id = 5, Name = "ABC"},
                new Person {Id = 77, Name = "DEF"},
                new Person {Id = 99, Name = null}
            };

            var file = await ExcelBuilder.BuildExcelFile(people);
            var path = GetExcelFilePath("PeopleExcel.xlsx");

            await File.WriteAllBytesAsync(path, file);
        }

        [Fact]
        public async Task BuildExcelFile_WhenValidStringListIsGiven_ShouldBuildExcelFile()
        {
            var textList = new List<string>
            {
                "TEXT 1",
                "TEXT 2",
                "TEXT 3"
            };

            var file = await ExcelBuilder.BuildExcelFile(textList);

            var path = GetExcelFilePath("StringsExcel.xlsx");

            await File.WriteAllBytesAsync(path, file);
        }

        [Fact]
        public async Task BuildExcelFile_WhenDynamicListIsGiven_ShouldBuildExcelFile()
        {
            var textList = new List<dynamic>
            {
                new Person {Id = 12345, Name = "Dogac"},
                "Dogac2",
                34567
            };

            var file = await ExcelBuilder.BuildExcelFile(textList);

            var path = GetExcelFilePath("DynamicExcel.xlsx");

            await File.WriteAllBytesAsync(path, file);
        }

        [Fact]
        public async Task BuildExcelFile_WhenDoubleListIsGiven_ShouldBuildExcelFile()
        {
            var textList = new List<double>
            {
                3.4345,
                1.3446889,
                45894.435
            };

            var file = await ExcelBuilder.BuildExcelFile(textList);

            var path = GetExcelFilePath("DoubleExcel.xlsx");

            await File.WriteAllBytesAsync(path, file);
        }
    }
}
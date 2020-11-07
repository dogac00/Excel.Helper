using System.IO;
using System.Linq;
using System.Threading.Tasks;
using AutoFixture;
using Xunit;

namespace Excel.Helper.Tests
{
    public class OverallTests
    {
        private readonly IFixture _fixture;

        public OverallTests()
        {
            _fixture = new Fixture();
        }

        [Fact]
        public async Task ShouldWorkWithByteArrayInput()
        {
            var people = _fixture
                .CreateMany<Person>()
                .ToList();
            var excel = await ExcelBuilder.BuildExcelFile(people);

            var peopleRead = await ExcelReader.ReadExcelFile<Person>(excel);

            Assert.Equal(people, peopleRead);
        }
        
        [Fact]
        public async Task ExcelFileShouldWorkWithStreamInput()
        {
            var people = _fixture
                .CreateMany<Person>()
                .ToList();
            var excel = await ExcelBuilder.BuildExcelFile(people);
            
            var fileName = "tempFile.xlsx";
            
            await File.WriteAllBytesAsync(fileName, excel);
            
            await using var fs = File.OpenRead(fileName);

            var peopleRead = await ExcelReader.ReadExcelFile<Person>(fs);

            Assert.Equal(people, peopleRead);
        }
        
        [Fact]
        public async Task ExcelFileShouldWorkWithFilePathInput()
        {
            var people = _fixture
                .CreateMany<Person>()
                .ToList();
            var excel = await ExcelBuilder.BuildExcelFile(people);
            
            var tempFilePath = "tempFile.xlsx";
            
            await File.WriteAllBytesAsync(tempFilePath, excel);

            var peopleRead = await ExcelReader.ReadExcelFile<Person>(tempFilePath);

            Assert.Equal(people, peopleRead);
        }
    }
}
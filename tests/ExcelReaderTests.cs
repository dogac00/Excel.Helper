using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Excel.Helper.Tests.Types;
using Xunit;
using static Excel.Helper.Tests.ExcelFilesUtil;

namespace Excel.Helper.Tests
{
    public class ExcelReaderTests
    {
        [Fact]
        public async Task ReadExcelFile_WhenEmptyStringCellsSupplied_ShouldReadCellsAsEmptyString()
        {
            var file = await ExcelBuilder.BuildExcelFile(new List<Person>
            {
                new Person() {Id = 12, Name = null},
                new Person() {Id = 15, Name = ""},
            });

            var readExcelFile = await ExcelReader.ReadExcelFile<Person>(file);

            foreach (var person in readExcelFile)
            {
                Assert.Equal("", person.Name);
            }
        }

        [Fact]
        public async Task ReadExcelFile_WhenValuesSupplied_ShouldRoundValuesOrConvertToStringRepresentation()
        {
            var invalidPeople = new List<PersonDoubleIdIntName>()
            {
                new PersonDoubleIdIntName() {Id = 12.5656, Name = 234},
                new PersonDoubleIdIntName() {Id = 15.134368, Name = 906}
            };
            var file = await ExcelBuilder.BuildExcelFile(invalidPeople);

            var peopleRead = await ExcelReader.ReadExcelFile<Person>(file);

            for (int i = 0; i < invalidPeople.Count; i++)
            {
                Assert.Equal(Math.Round(invalidPeople[i].Id), peopleRead[i].Id);
                Assert.Equal(invalidPeople[i].Name.ToString(), peopleRead[i].Name);
            }
        }

        [Fact]
        public async Task ReadExcelFile_WhenInvalidPeopleExcelFileIsGiven_ShouldThrowInvalidExcelException()
        {
            var invalidPeople = new List<PersonStringIdIntName>
            {
                new PersonStringIdIntName() {Id = "12.5656", Name = 234},
                new PersonStringIdIntName() {Id = "15.134368", Name = 906}
            };
            var file = await ExcelBuilder.BuildExcelFile(invalidPeople);

            await Assert.ThrowsAsync<InvalidExcelException>(async () =>
            {
                var peopleRead = await ExcelReader.ReadExcelFile<Person>(file);
            });
        }

        [Fact]
        public async Task ReadExcelFile_WhenInvalidDoubleExcelFileIsGiven_ShouldThrowInvalidExcelException()
        {
            var file = GetExcelsFolderPath("InvalidDoubleExcel.xlsx");

            await Assert.ThrowsAsync<InvalidExcelException>(async () =>
            {
                await ExcelReader.ReadExcelFile<double>(file);
            });
        }
        
        [Fact]
        public async Task ReadExcelFile_WhenPeopleExcelWithEmptyDoubleCell_ShouldThrowInvalidExcelException()
        {
            var file = GetExcelsFolderPath("PeopleExcelWithEmptyDoubleCell.xlsx");

            var exception = await Assert.ThrowsAsync<InvalidExcelException>(async () =>
            {
                await ExcelReader.ReadExcelFile<PersonIntIdStringNameDoubleTest>(file);
            });
            
            Assert.Single(exception.InvalidColumns);
        }
    }
}
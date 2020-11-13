using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Xunit;

namespace Excel.Helper.Tests
{
    public class ExcelReaderTests
    {
        private string GetExcelsFolderPath(string fileName)
        {
            return $"../../../Excels/{fileName}";
        }

        [Fact]
        public async Task ReadExcel_ShouldWorkWithEmptyCells()
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
        public async Task ReadExcel_ShouldWorkWithInvalidCells()
        {
            var invalidPeople = new List<InvalidPerson>()
            {
                new InvalidPerson() {Id = 12.5656, Name = 234},
                new InvalidPerson() {Id = 15.134368, Name = 906}
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
        public async Task ReadExcel_ShouldThrowWhenCannotConvert()
        {
            var invalidPeople = new List<InvalidPerson2>
            {
                new InvalidPerson2() {Id = "12.5656", Name = 234},
                new InvalidPerson2() {Id = "15.134368", Name = 906}
            };
            var file = await ExcelBuilder.BuildExcelFile(invalidPeople);

            await Assert.ThrowsAsync<InvalidExcelException>(async () =>
            {
                var peopleRead = await ExcelReader.ReadExcelFile<Person>(file);
            });
        }

        [Fact]
        public async Task ReadExcel_InvalidDoubleExcel_ShouldConvert()
        {
            var file = GetExcelsFolderPath("InvalidDoubleExcel.xlsx");

            await Assert.ThrowsAsync<InvalidExcelException>(async () =>
            {
                await ExcelReader.ReadExcelFile<double>(file);
            });
        }
    }

    public class InvalidPerson
    {
        public double Id { get; set; }
        public int Name { get; set; }
    }

    public class InvalidPerson2
    {
        public string Id { get; set; }
        public int Name { get; set; }
    }
}
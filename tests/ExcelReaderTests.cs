using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using Excel.Helper.Tests.Types;
using FluentAssertions;
using Microsoft.VisualStudio.TestPlatform.ObjectModel;
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
        public async Task ReadExcelFile_WhenPeopleExcelIsSubmitted_ShouldReadPeopleExcel()
        {
            var file = GetExcelFilePath("PeopleExcel.xlsx");

            var peopleRead = await ExcelReader.ReadExcelFile<Person>(file);

            var people1 = peopleRead[0];
            var people2 = peopleRead[1];
            var people3 = peopleRead[2];

            Assert.Equal(5, people1.Id);
            Assert.Equal("ABC", people1.Name);
            Assert.Equal(77, people2.Id);
            Assert.Equal("DEF", people2.Name);
            Assert.Equal(99, people3.Id);
            Assert.Equal("", people3.Name);
        }

        [Fact]
        public async Task ReadExcelFile_WhenPeopleExcelSuppliedAndParsedToDoubleList_ShouldReadFirstColumn()
        {
            var file = GetExcelFilePath("PeopleExcel.xlsx");

            var doubles = await ExcelReader.ReadExcelFile<double>(file);

            doubles[0].Should().Be(5);
            doubles[1].Should().Be(77);
            doubles[2].Should().Be(99);
        }

        [Fact]
        public async Task ReadExcelFile_WhenDoubleExcelWithEmptyColumnSupplied_ShouldNotParseEmptyCell()
        {
            var file = GetExcelFilePath("DoubleExcelWithEmptyColumn.xlsx");

            var doubles = await ExcelReader.ReadExcelFile<double>(file);

            doubles.Count.Should().Be(3);
            doubles.First().Should().Be(1.32);
            doubles[1].Should().Be(3.45);
            doubles[2].Should().Be(64564767);
        }

        [Fact]
        public async Task ReadExcelFile_ShouldParseStringExcel()
        {
            var file = GetExcelFilePath("DynamicExcel.xlsx");

            var content = await ExcelReader.ReadExcelFile<string>(file);

            content.First().Should().Be("Id = 12345, Name = Dogac");
            content.Second().Should().Be("Dogac2");
            content.Third().Should().Be("34567");
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
            var file = GetExcelFilePath("InvalidDoubleExcel.xlsx");

            var exception = await Assert.ThrowsAsync<InvalidExcelException>(async () =>
            {
                await ExcelReader.ReadExcelFile<double>(file);
            });

            exception.InvalidColumns.Count.Should().Be(1);
            var first = exception.InvalidColumns.First();
            
            first.Column.Should().Be(1);
            first.Row.Should().Be(2);
            first.Value.Should().Be("Dogac");
            first.ConversionType.Should().Be(typeof(double));
        }

        [Fact]
        public async Task ReadExcelFile_CountShouldBeThree()
        {
            var file = GetExcelFilePath("InvalidDoubleExcel.xlsx");

            var result = await ExcelReader.ReadExcelFile<string>(file);

            result.Count.Should().Be(3);
        }

        [Fact]
        public async Task ReadExcelFile_WhenPeopleExcelWithEmptyDoubleCell_ShouldThrowInvalidExcelException()
        {
            var file = GetExcelFilePath("PeopleExcelWithEmptyDoubleCell.xlsx");

            var exception = await Assert.ThrowsAsync<InvalidExcelException>(async () =>
            {
                await ExcelReader.ReadExcelFile<PersonIntIdStringNameDoubleTest>(file);
            });

            Assert.Single(exception.InvalidColumns);
        }
    }

    static class EnumerableExtensions
    {
        public static T Second<T>(this IEnumerable<T> source)
        {
            if (source == null)
            {
                throw new ArgumentNullException(nameof(source));
            }

            using var enumerator = source.GetEnumerator();

            if (!enumerator.MoveNext() |
                !enumerator.MoveNext())
            {
                throw new ArgumentException("Not enough elements in source.");
            }

            return enumerator.Current;
        }

        public static T Third<T>(this IEnumerable<T> source)
        {
            if (source == null)
            {
                throw new ArgumentNullException(nameof(source));
            }

            using var enumerator = source.GetEnumerator();

            if (!enumerator.MoveNext() |
                !enumerator.MoveNext() |
                !enumerator.MoveNext())
            {
                throw new ArgumentException("Not enough elements in source.");
            }

            return enumerator.Current;
        }

        public static T Fourth<T>(this IEnumerable<T> source)
        {
            if (source == null)
            {
                throw new ArgumentNullException(nameof(source));
            }

            using var enumerator = source.GetEnumerator();

            if (!enumerator.MoveNext() |
                !enumerator.MoveNext() |
                !enumerator.MoveNext() |
                !enumerator.MoveNext())
            {
                throw new ArgumentException("Not enough elements in source.");
            }

            return enumerator.Current;
        }
    }
}
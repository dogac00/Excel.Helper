using System;
using System.Linq;
using System.Threading.Tasks;
using AutoFixture;
using Xunit;

namespace Excel.Helper.Tests
{
    public class ExcelBuilderTests
    {
        private readonly IFixture _fixture;

        public ExcelBuilderTests()
        {
            _fixture = new Fixture();
        }

        [Fact]
        public async Task BuildExcelFile_WhenInvalidListIsGiven_ShouldThrow()
        {
            await Assert.ThrowsAsync<ArgumentNullException>(async () =>
            {
                await ExcelBuilder.BuildExcelFile((int[]) null);
            });
        }
        
        [Fact]
        public async Task BuildExcelFile_WhenInvalidNameIsGiven_ShouldThrow()
        {
            await Assert.ThrowsAsync<ArgumentException>(async () =>
            {
                var list = _fixture
                    .CreateMany<string>()
                    .ToList();
                await ExcelBuilder.BuildExcelFile(list, null);
            });
        }
    }
}
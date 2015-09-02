using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using FluentAssertions;

namespace ExcelParser.Tests
{
    [TestClass]
    public class SimpleExcelParser_MultiSheetWorkbookTests
    {
        SimpleExcelParser parser;
        string fileName;
        [TestInitialize]
        public void Setup()
        {
            parser = new SimpleExcelParser();
            fileName = "..\\..\\..\\MultiSheetWorkbook.xlsx";
        }

        [TestMethod]
        public void Parse_Should_Return_One_Objects()
        {
            // Arrange


            // Act
            var results = parser.Parse<Holiday>(fileName, "Sheet2");

            // Assert
            results.Should().HaveCount(1);
        }


        [TestMethod]
        public void Parse_Should_Return_Easter_As_First_Object()
        {
            // Arrange


            // Act
            var results = parser.Parse<Holiday>(fileName, "Sheet2");

            // Assert
            var christmas = results.First();
            christmas.Name.Should().Be("Easter");
            christmas.NumberOfDays.Should().Be(1);
            christmas.StartDate.Should().Be(new DateTime(2015, 1, 1));
            christmas.OvertimeRate.Should().Be(150m);
            christmas.SalaryBonus.Should().BeApproximately(0.003846154d, 0.000000001d);
            christmas.UniqueId.Should().Be("AAAAAAAA-8BB3-413F-BBBE-CCE71E470594");
        }

        public class Holiday
        {
            public string Name { get; set; }
            public int NumberOfDays { get; set; }
            public DateTime StartDate { get; set; }
            public Decimal OvertimeRate { get; set; }
            public Double SalaryBonus { get; set; }
            public Guid UniqueId { get; set; }
        }
    }
}

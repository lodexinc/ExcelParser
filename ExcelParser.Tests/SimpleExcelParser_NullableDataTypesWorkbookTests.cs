using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using FluentAssertions;

namespace ExcelParser.Tests
{
    [TestClass]
    public class SimpleExcelParser_NullableDataTypesWorkbookTests
    {
        SimpleExcelParser parser;
        string fileName;
        [TestInitialize]
        public void Setup()
        {
            parser = new SimpleExcelParser();
            fileName = "..\\..\\..\\NullableDataTypesWorkbook.xlsx";
        }

        [TestMethod]
        public void Parse_Should_Return_Two_Objects()
        {
            // Arrange


            // Act
            var results = parser.Parse<Holiday>(fileName);

            // Assert
            results.Should().HaveCount(2);
        }


        [TestMethod]
        public void Parse_Should_Return_Christmas_As_First_Object()
        {
            // Arrange


            // Act
            var results = parser.Parse<Holiday>(fileName);

            // Assert
            var christmas = results.First();
            christmas.Name.Should().Be("Christmas");
            christmas.NumberOfDays.Should().Be(2);
            christmas.StartDate.Should().Be(new DateTime(2015, 12, 25));
            christmas.OvertimeRate.Should().Be(150m);
            christmas.SalaryBonus.Should().BeApproximately(0.007692308d, 0.000000001d);
            christmas.UniqueId.Should().Be("A23B58EC-8BB3-413F-BBBE-CCE71E470594");
        }


        [TestMethod]
        public void Parse_Should_Return_NULLs_For_Second_Object()
        {
            // Arrange


            // Act
            var results = parser.Parse<Holiday>(fileName);

            // Assert
            var obj = results.Skip(1).First();
            obj.Name.Should().BeNull();
            obj.NumberOfDays.Should().NotHaveValue();
            obj.StartDate.Should().NotHaveValue();
            obj.OvertimeRate.Should().NotHaveValue();
            obj.SalaryBonus.Should().NotHaveValue();
            obj.UniqueId.Should().NotHaveValue();
        }

        public class Holiday
        {
            public string Name { get; set; }
            public int? NumberOfDays { get; set; }
            public DateTime? StartDate { get; set; }
            public Decimal? OvertimeRate { get; set; }
            public Double? SalaryBonus { get; set; }
            public Guid? UniqueId { get; set; }
        }
    }
}

using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using FluentAssertions;
using System.Linq;

namespace ExcelParser.Tests
{
    [TestClass]
    public class SimpleExcelParser_SimpleWorkbookTests
    {
        SimpleExcelParser parser;
        string fileName;
        [TestInitialize]
        public void Setup()
        {
            parser = new SimpleExcelParser();
            fileName = "..\\..\\..\\SimpleWorkbook.xlsx";
        }

        [TestMethod]
        public void Parse_Should_Return_6_Objects()
        {
            // Arrange

            // Act
            var results = parser.Parse<Person>(fileName);

            // Assert
            results.Should().HaveCount(6);
        }

        [TestMethod]
        public void Parse_Should_Return_Tim_As_First_Object()
        {
            // Arrange

            // Act
            var results = parser.Parse<Person>(fileName);

            // Assert
            var tim = results.First();
            tim.FirstName.Should().Be("Timothy");
            tim.MiddleName.Should().Be("John");
            tim.LastName.Should().Be("Rayburn");
        }

        [TestMethod]
        public void Parse_Should_Return_Tim_As_Fourth_Object()
        {
            // Arrange

            // Act
            var results = parser.Parse<Person>(fileName);

            // Assert
            var tim = results.Skip(3).First();
            tim.FirstName.Should().Be("George");
            tim.MiddleName.Should().BeNull();
            tim.LastName.Should().Be("Washington");
        }

        public class Person
        {
            public string FirstName { get; set; }
            public string MiddleName { get; set; }
            public string LastName { get; set; }
        }
    }
}

using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using FluentAssertions;

namespace ExcelParser.Tests
{
    [TestClass]
    public class DehumanizeNameInterceptor_Tests
    {
        private DehumanizeNameInterceptor target;

        [TestInitialize]
        public void Setup()
        {
            target = new DehumanizeNameInterceptor();
        }

        [TestMethod]
        public void Intercept_Should_ProperCase_Names()
        {
            // Arrange

            // Act
            var result = target.Intercept("this is awesome");

            // Assert
            result.Should().Be("ThisIsAwesome");
        }

        [TestMethod]
        public void Order_Should_Default_To_Zero()
        {
            // Arrange


            // Act
            var result = target.Order;

            // Assert
            result.Should().Be(0);
        }

        [TestMethod]
        public void Order_Should_Set_From_Constructor()
        {
            // Arrange
            target = new DehumanizeNameInterceptor(100);

            // Act
            var result = target.Order;

            // Assert
            result.Should().Be(100);
        }
    }
}

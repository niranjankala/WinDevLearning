using Calculator.Library;
using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Calculator.Library.Tests
{
    [TestClass]
    public class CalculatorTests
    {
        [TestMethod]
        public void Divide_PositiveNumbers_ReturnsPositiveQuotient()
        {
            //Arrange
            int expected = 5;
            int numerator = 20;
            int denominator = 4;

            //Act
            int actual = Calculator.Divide(numerator, denominator);

            //Assert
            Assert.AreEqual(expected, actual);
        }
        [TestMethod]
        public void Divide_PositiveNumeratorAndNegativeDenominator_ReturnsNegativeQuotient()
        {
            //Arrange
            int expected = -5;
            int numerator = 20;
            int denominator = -4;

            //Act
            int actual = Calculator.Divide(numerator, denominator);

            //Assert
            Assert.AreEqual(expected, actual);
        }
        [TestMethod]
        public void Divide_NegativeNumbers_ReturnsPositiveQuotient()
        {
            //Arrange
            int expected = 5;
            int numerator = -20;
            int denominator = -4;

            //Act
            int actual = Calculator.Divide(numerator, denominator);

            //Assert
            Assert.AreEqual(expected, actual);
        }
        [TestMethod]
        [ExpectedException(typeof(DivideByZeroException))]
        public void Divide_DenominatorIsZero_ThrowDivideByZeroException()
        {
            //Arrange
            int numerator = 020;
            int denominator = 0;

            //Act
            try
            {
                int actual = Calculator.Divide(numerator, denominator);
            }
            catch (Exception ex)
            {
                //Assert
                Assert.AreEqual("Denominator cannot be ZERO", ex.Message);
                throw;
            }
        }

    }
}

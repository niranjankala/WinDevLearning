using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculator.Library
{
    public class Calculator
    {
        
        public static int Divide(int numerator, int denominator)
        {
            if (denominator == 0)
                throw new DivideByZeroException("Denominator cannot be ZERO");
            int result = numerator / denominator;
            return result;
        }

        public static int Add(int firstNumber, int secondNumber)
        {
            if (IsPositive(firstNumber) && IsPositive(secondNumber))
            {
                int result = firstNumber + secondNumber;
                return result;
            }
            else
            {
                throw new ArgumentException("Only positive numbers are allowed");
            }
        }
        private static bool IsPositive(int number)
        {
            return number>0;
        }
    }
}

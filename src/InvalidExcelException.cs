using System;

namespace Excel.Helper
{
    public class InvalidExcelException : Exception
    {
        public InvalidExcelException(string message) : base(message)
        {
        }
    }
}
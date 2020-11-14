using System;
using System.Collections.Generic;

namespace Excel.Helper
{
    public class InvalidExcelException : Exception
    {
        public List<InvalidColumn> InvalidColumns { get; }
        
        public InvalidExcelException(List<InvalidColumn> invalidColumns)
        {
            this.InvalidColumns = invalidColumns;
        }
    }
}
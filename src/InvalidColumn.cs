using System;

namespace Excel.Helper
{
    public class InvalidColumn
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public object Value { get; set; }
        public FormatException Exception { get; set; }
        public Type ConversionType { get; set; }
    }
}
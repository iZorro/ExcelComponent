using System;

namespace ExcelComponent.Excel
{
    internal static class Util
    {
        internal static void ValidateUShort(int theInt, string fieldName)
        {
            if (theInt < ushort.MinValue || theInt > ushort.MaxValue)
                throw new Exception (string.Format("{0} value {1} must be between 1 and {2}", fieldName, theInt, ushort.MaxValue - 1));
        }
    }
}

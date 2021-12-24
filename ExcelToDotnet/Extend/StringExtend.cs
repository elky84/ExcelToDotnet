using System.Text.RegularExpressions;

namespace ExcelToDotnet.Extend
{
    public static class StringExtend
    {
        public static string RemoveSpecialCharacters(this string str)
        {
            return Regex.Replace(str, "[^a-zA-Z0-9_.]+", "", RegexOptions.Compiled);
        }

        public static string ExtractDataTypeInList(this string str)
        {
            return str.Split('<', '>')[1];
        }

        public static string ToStringValue(this string? str)
        {
            return (str as object).ToStringValue();
        }

        public static string ToSafeString(this string? str)
        {
            return (str as object).ToSafeString();
        }

        public static double ToDoubleValue(this string str)
        {
            return Convert.ToDouble(str);
        }

        public static string ToCamelCase(this string str)
        {
            return Char.ToUpperInvariant(str[0]) + str.Substring(1).ToLower();
        }
    }
}

namespace ExcelToDotnet.Extend
{
    public static class ObjectExtend
    {
        public static string ToStringValue(this object? obj)
        {
            if (obj == null)
            {
                return string.Empty;
            }
            else if (obj.GetType() == typeof(string))
            {
                return (string)obj;
            }
            else if (obj.GetType().IsPrimitive || obj.GetType() == typeof(DateTime) || obj.GetType() == typeof(TimeSpan))
            {
                return ToSafeString(obj);
            }
            else
            {
                return string.Empty;
            }
        }

        public static string ToSafeString(this object? obj)
        {
            if (obj == null)
            {
                return "";
            }

            var str = obj.ToString();
            if (str == null)
            {
                return "";
            }

            return (string)str;
        }

        public static double ToDoubleValue(this object obj)
        {
            return Convert.ToDouble(obj);
        }
    }
}

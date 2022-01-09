namespace ExcelToDotnet.Extend
{
    public static class ConvertExtend
    {
        public static string ToMemberDefinition<TKey, TValue>(this IDictionary<TKey, TValue> dictionary)
        {
            return "{" + string.Join(",", dictionary.Select(kv => kv.Key + "=" + kv.Value).ToArray()) + "}";
        }


        public static List<KeyValuePair<string, string>> Convert(this List<string?> dataTypes, bool nullable)
        {
            return dataTypes.ConvertAll(x =>
            {
                if (x == null)
                {
                    return new KeyValuePair<string, string>("", "");
                }

                if (x.GetType() != typeof(string))
                {
                    return new KeyValuePair<string, string>(x.ToSafeString(), x.ToSafeString());
                }

                var str = (string)x;
                if (str.StartsWith("$"))
                {
                    return new KeyValuePair<string, string>(str, nullable ? "string?" : "string");
                }
                if (str.StartsWith("List") && str.Contains('$'))
                {
                    return new KeyValuePair<string, string>(str, nullable ? "List<string>?" : "List<string>");
                }
                if (str.StartsWith("~"))
                {
                    return new KeyValuePair<string, string>(str, "int");
                }
                if (str.StartsWith("%"))
                {
                    return new KeyValuePair<string, string>(str, "double");
                }
                if (str.StartsWith("!"))
                {
                    return new KeyValuePair<string, string>("!", str.Replace("!", string.Empty));
                }
                if (str.StartsWith("*"))
                {
                    return new KeyValuePair<string, string>("*", str.Replace("*", string.Empty));
                }
                return new KeyValuePair<string, string>(str, str);
            });
        }

        public static List<KeyValuePair<int, string>> ConvertToReferenceId(this List<string?> dataTypes)
        {
            var list = new List<KeyValuePair<int, string>>();
            for (int x = 0; x < dataTypes.Count; ++x)
            {
                var value = dataTypes[x].ToStringValue();
                if (value.StartsWith("$") || (value.StartsWith("List") && value.Contains("$")))
                {
                    list.Add(new KeyValuePair<int, string>(x, value));
                }
            }
            return list;
        }

        public static List<KeyValuePair<int, string>> ConvertToSubIndex(this List<string?> dataTypes)
        {
            var list = new List<KeyValuePair<int, string>>();
            for (int x = 0; x < dataTypes.Count; ++x)
            {
                var value = dataTypes[x].ToStringValue();
                if (value.StartsWith("~"))
                {
                    list.Add(new KeyValuePair<int, string>(x, value));
                }
            }
            return list;
        }

        public static List<KeyValuePair<int, string>> ConvertToProbability(this List<string?> dataTypes)
        {
            var list = new List<KeyValuePair<int, string>>();
            for (int x = 0; x < dataTypes.Count; ++x)
            {
                var value = dataTypes[x].ToStringValue();
                if (value.StartsWith("%"))
                {
                    list.Add(new KeyValuePair<int, string>(x, value));
                }
            }
            return list;
        }


        public static List<KeyValuePair<int, string>> ConvertToReferenceEnum(this List<string?> dataTypes)
        {
            var list = new List<KeyValuePair<int, string>>();
            for (int x = 0; x < dataTypes.Count; ++x)
            {
                var value = dataTypes[x].ToStringValue();
                if ((value.StartsWith("List") && (value.EndsWith("Type>?") || value.EndsWith("Type>"))) || 
                    value.EndsWith("Type") || 
                    value.EndsWith("Type?"))
                {
                    list.Add(new KeyValuePair<int, string>(x, value));
                }
            }
            return list;
        }


    }
}

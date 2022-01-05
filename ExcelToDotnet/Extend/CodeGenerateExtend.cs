using System.Data;

namespace ExcelToDotnet.Extend
{
    public static class CodeGenerateExtend
    {
        public static void GenerateCode(this List<string> strings, StreamWriter outputFile, DataTable dt, List<string?> dataTypes, Dictionary<string, string> patterns, bool nullable)
        {
            var dataTypesConverted = dataTypes.Convert(nullable);
            foreach (var pattern in patterns)
            {
                int index = strings.FindIndex(str => str.Contains(pattern.Key));
                if (index != -1)
                {
                    strings[index] = strings[index].Replace(pattern.Key, pattern.Value);
                }
            }

            int insertIndex = strings.FindIndex(str => str.Contains("##INSERT##"));
            strings.RemoveRange(insertIndex, 1);

            for (int n = dt.Columns.Count - 1; n >= 0; --n)
            {
                var plainDataType = dataTypesConverted[n].Key;
                var dataType = dataTypesConverted[n].Value;

                strings.Insert(insertIndex, string.Empty);

                strings.Insert(insertIndex, string.Format("        public {0} {1}", dataType, dt.Columns[n].ColumnName) + " { get; set; }");
                if (dataType.EndsWith("Type") || dataType.EndsWith("Type?"))
                {
                    strings.Insert(insertIndex, string.Format($"        [JsonConverter(typeof(JsonEnumConverter<{dataType.RemoveSpecialCharacters()}>))]"));
                }
                else if (dataType.StartsWith("List") && (dataType.EndsWith("Type>") || dataType.EndsWith("Type>?")))
                {
                    strings.Insert(insertIndex, string.Format($"        [JsonConverter(typeof(JsonEnumsConverter<{dataType.ExtractDataTypeInList()}>))]"));
                }
            }

            foreach (var str in strings)
            {
                outputFile.WriteLine(str);
            }
        }

        public static void GenerateEnumCode(this List<string> strings, StreamWriter outputFile, DataTable dt, Dictionary<string, string> patterns)
        {
            foreach (var pattern in patterns)
            {
                int index = strings.FindIndex(str => str.Contains(pattern.Key));
                if (index != -1)
                {
                    strings[index] = strings[index].Replace(pattern.Key, pattern.Value);
                }
            }

            int insertIndex = strings.FindIndex(str => str.Contains("##INSERT##"));
            strings.RemoveRange(insertIndex, 1);

            for (int n = dt.Rows.Count - 1; n >= 0; --n)
            {
                strings.Insert(insertIndex, string.Empty);
                strings.Insert(insertIndex, string.Format("        {0},", dt.Rows[n].ItemArray[0]));
                strings.Insert(insertIndex, string.Format("        [Description(\"{0}\")]", dt.Rows[n].ItemArray[1]));
            }

            foreach (var str in strings)
            {
                outputFile.WriteLine(str);
            }
        }
    }
}

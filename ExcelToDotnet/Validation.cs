using ExcelToDotnet.Extend;
using Newtonsoft.Json;
using System.CodeDom.Compiler;
using System.Data;
using System.Drawing;
using System.Text.RegularExpressions;

namespace ExcelToDotnet
{
    public static class Validation
    {
        public static void Failed(string format, params object[] args)
        {
            LogExtend.Color(ConsoleColor.Red, ConsoleColor.Black, format, args);
            Environment.ExitCode = ErrorCode.VALIDATION_FAILED;
        }

        private static HashSet<string> ToHashSet(this List<Dictionary<string, object>> list, string tableName, string key, bool checkDuplicate)
        {
            var hashSet = new HashSet<string>();
            foreach (var dic in list)
            {
                if (!dic.ContainsKey(key))
                {
                    Failed($"키 참조에 실패했습니다. JSON에 존재하지 않는 키를 참조 값으로 사용했습니다. <Table:{tableName}, Key:{key}>");
                    continue;
                }

                if (checkDuplicate)
                {
                    if (hashSet.Contains(dic[key]))
                    {
                        Failed($"중복된 Enum을 사용했습니다. <Table:{tableName}, Key:{key}>");
                    }
                }

                if (!hashSet.Contains(dic[key]))
                {
                    hashSet.Add(dic[key].ToStringValue());
                }
            }

            return hashSet;
        }

        public static void ValidateValue(this DataTable dt, string keyword, List<string?> dataTypes)
        {
            for (int n = 0; n < dt.Rows.Count; ++n)
            {
                for (int x = 0; x < dataTypes.Count; ++x)
                {
                    if (dataTypes[x] == null)
                        continue;

                    var dataTypeString = dataTypes[x].ToStringValue();
                    var obj = dt.Rows[n].ItemArray.ToArray()[x];
                    if (obj == null)
                    {
                        continue;
                    }

                    var str = obj.ToStringValue();
                    var itemArray = dt.Rows[n].ItemArray;
                    if (itemArray == null)
                    {
                        continue;
                    }

                    if (IsPrimitiveType(dataTypeString))
                    {
                        try
                        {
                            if (false == ValidationConvert(dataTypeString, str))
                            {
                                throw new Exception("PrimitiveType Data Convert Failed");
                            }
                        }
                        catch (Exception e)
                        {
                            Failed($"[{keyword}] {dataTypeString} 타입으로의 변환에 실패했습니다. 유효하지 않은 값이 들어있을 확률이 높습니다. <Table:{dt.TableName}, Value:{str}, Id:{itemArray.ToIdString(dt)}. Exception:{e.Message}>");
                        }
                        continue;
                    }
                    else
                    {
                        if (!dataTypeString.Contains("?") && string.IsNullOrEmpty(str) && !dataTypeString.StartsWith("List"))
                        {
                            Failed($"[{keyword}] {dataTypeString} 타입에는 null이 허용되지 않습니다. <Table:{dt.TableName}, Key:{str}, Id:{itemArray.ToIdString(dt)}>");
                        }
                    }
                }
            }
        }

        public static void ValidateSubIndex(this DataTable dt, string keyword, List<string?> dataTypes)
        {
            var subIndexes = dataTypes.ConvertToSubIndex();
            foreach (var subIndex in subIndexes)
            {
                var index = subIndex.Key;
                var columnName = dt.Columns[index].ColumnName;

                string refDataType = subIndex.Value.RemoveSpecialCharacters();
                var refColumn = dt.Columns.Cast<DataColumn>().FirstOrDefault(x => x.ColumnName == refDataType);
                var refIndex = dt.Columns.IndexOf(refColumn);

                var path = string.Format($"output/{keyword}/{dt.TableName.RemoveSpecialCharacters()}.json");
                if (!File.Exists(path))
                {
                    Failed($"[{keyword}] JSON 파일이 존재하지 않습니다. 없는 테이블을 참조했습니다. " +
                        $"<Table:{dt.TableName}, 참조 컬럼:{refColumn}, 참조데이터타입:{refDataType}, " +
                        $"Column:{dt.Columns[index].ColumnName}, Path:{path}>");
                    continue;
                }

                var json = File.ReadAllText(path);
                var list = JsonConvert.DeserializeObject<List<Dictionary<string, object>>>(json);

                for (int n = 0; n < dt.Rows.Count; ++n)
                {
                    var itemArray = dt.Rows[n].ItemArray.ToArray();
                    if (itemArray == null)
                        continue;

                    var item = itemArray[index];
                    if (item == null)
                        continue;

                    var str = item.ToStringValue();

                    if (string.IsNullOrEmpty(str))
                    {
                        Failed($"[{keyword}] 내부 참조 과정에서 Null이 허용되지 않는데, 비어있는 데이터를 발견했습니다. " +
                            $"<Table:{dt.TableName}, 컬럼:{dt.Columns[index].ColumnName} " +
                            $"참조 컬럼:{refColumn}, 참조데이터타입:{refDataType}, Key:{str}, Id:{itemArray.ToIdString(dt)}>");
                    }
                }

                var groupByRow = dt.Rows.Cast<DataRow>().GroupBy(x =>
                {
                    var left = x.ItemArray.ToArray()[index];
                    var right = x.ItemArray.ToArray()[refIndex];
                    if (left == null ||
                        right == null)
                        return new KeyValuePair<string, string>(string.Empty, string.Empty);

                    return new KeyValuePair<string, string>(left.ToStringValue(), right.ToStringValue());
                });

                foreach (var row in groupByRow.Where(g => !string.IsNullOrEmpty(g.Key.Key) && g.Count() > 1)
                                              .Select(y => y.Key)
                                              .ToList())
                {
                    Failed($"[{keyword}] 내부 참조 과정에서 중복된 행이 발견되었습니다. " +
                        $"<Table:{dt.TableName}, 참조 컬럼:{refColumn}, 참조데이터타입:{refDataType}, " +
                        $"Column:{columnName}, Row:{row}>");
                }
            }
        }

        public static void ValidationProbability(this DataTable dt, string keyword, List<string?> dataTypes)
        {
            var probabilities = dataTypes.ConvertToProbability();
            foreach (var probability in probabilities)
            {
                var index = probability.Key;
                var columnName = dt.Columns[index].ColumnName;

                string refDataType = probability.Value.RemoveSpecialCharacters();
                var refColumn = dt.Columns.Cast<DataColumn>().FirstOrDefault(x => x.ColumnName == columnName);
                var refIndex = dt.Columns.IndexOf(refColumn);

                var probRefColumnName = probability.Value.RemoveSpecialCharacters();
                var probRefColumn = dt.Columns.Cast<DataColumn>().FirstOrDefault(x => x.ColumnName == probRefColumnName);
                var probRefIndex = dt.Columns.IndexOf(probRefColumn);

                var path = string.Format($"output/{keyword}/{dt.TableName.RemoveSpecialCharacters()}.json");
                if (!File.Exists(path))
                {
                    Failed($"[{keyword}] JSON 파일이 존재하지 않습니다. 없는 테이블을 참조했습니다. " +
                        $"<Table:{dt.TableName}, 참조 컬럼:{refColumn}, 참조데이터타입:{refDataType}, " +
                        $"Column:{dt.Columns[index].ColumnName}, Path:{path}>");
                    continue;
                }

                var json = File.ReadAllText(path);
                var list = JsonConvert.DeserializeObject<List<Dictionary<string, object>>>(json);

                for (int n = 0; n < dt.Rows.Count; ++n)
                {
                    var item = dt.Rows[n].ItemArray.ToArray()[index];
                    if (item == null)
                    {
                        continue;
                    }

                    var str = item.ToStringValue();

                    if (string.IsNullOrEmpty(str))
                    {
                        Failed($"[{keyword}] 확률 참조 과정에서 Null이 허용되지 않는데, 비어있는 데이터를 발견했습니다. " +
                            $"<Table:{dt.TableName}, 컬럼:{dt.Columns[index].ColumnName}, 참조 컬럼:{refColumn}, " +
                            $"참조Id:{refDataType}, Key:{str}, Id:{dt.Rows[n].ItemArray.ToIdString(dt)}>");
                    }
                }

                var group = dt.Rows.Cast<DataRow>().GroupBy(x => x.ItemArray.ToArray()[probRefIndex]);

                foreach (var sum in group.Select(g => g.Sum(x => x.ItemArray.ToArray()[refIndex]!.ToDoubleValue())))
                {
                    if (Math.Abs(sum - 100.0) > 0.00000001)
                    {
                        Failed($"[{keyword}] 확률 참조 과정에서 합산 값이 100%가 아닌 행이 발견되었습니다. <Table:{dt.TableName}, 참조 컬럼:{refColumn}, " +
                            $"참조데이터타입:{refDataType}, Column:{columnName}, ProbRefColumn:{probRefColumnName}," +
                            $"Sum:{sum}>");
                    }

                }
            }
        }


        public static void ValidateReference(this DataTable dt, string keyword, List<string?> dataTypes)
        {
            var refIds = dataTypes.ConvertToReferenceId();
            foreach (var pair in refIds)
            {
                var index = pair.Key;

                if (pair.Value.StartsWith("List"))
                {
                    dt.ValidateReference(keyword, index, pair.Value.ExtractDataTypeInList(), true);
                }
                else
                {
                    dt.ValidateReference(keyword, index, pair.Value, false);
                }
            }
        }

        private static bool ValidateReference(this DataTable dt, string keyword, int index, string refId, bool isList)
        {
            var realRefId = refId.Contains(".") ? refId.Split(".")[0] : refId;
            var path = string.Format($"output/{keyword}/{realRefId.RemoveSpecialCharacters()}.json");
            if (!File.Exists(path))
            {
                Failed($"[{keyword}] JSON 파일이 존재하지 않습니다. 없는 테이블을 참조했습니다. <Table:{dt.TableName}, 참조Id:{refId}, Column:{dt.Columns[index].ColumnName}, Path:{path}>");
                return false;
            }

            var json = File.ReadAllText(path);
            var list = JsonConvert.DeserializeObject<List<Dictionary<string, object>>>(json);
            if (list == null)
            {
                Failed($"[{keyword}] JSON Deserialize에 실패했습니다. <Table:{dt.TableName}, 참조Id:{refId}, Column:{dt.Columns[index].ColumnName}, Path:{path}, Json:{json}>");
                return false;
            }

            var key = refId.Contains(".") ? refId.Split(".")[1] : "Id";
            var hashSet = list.ToHashSet(dt.TableName, key, true);

            for (int n = 0; n < dt.Rows.Count; ++n)
            {
                var obj = dt.Rows[n].ItemArray.ToArray()[index];
                if (obj == null)
                {
                    Failed($"[{keyword}] 유효하지 않은  Table:{dt.TableName}, <참조Id:{refId}, Column:{dt.Columns[index].ColumnName}, Path:{path}>");
                    return false;
                }

                var str = obj.ToStringValue();
                if (realRefId.EndsWith("?") && string.IsNullOrEmpty(str))
                {
                    continue;
                }

                if (isList)
                {
                    if (obj.GetType() != typeof(DBNull))
                    {
                        var strings = JsonConvert.DeserializeObject<List<string>>(str);
                        if (strings == null)
                        {
                            Failed($"[{keyword}] 문자열로 변환에 실패했습니다. 참조에 실패했습니다. <Table:{dt.TableName}>");
                            return false;
                        }

                        foreach (var checkKey in strings)
                        {
                            if (false == hashSet.Contains(checkKey))
                            {
                                Failed($"[{keyword}] 참조 테이블 내에서 유효한 값이 없습니다. <Table:{dt.TableName}, 컬럼:{dt.Columns[index].ColumnName} " +
                                    $"참조Id:{refId}, Key:{checkKey}, Id:{dt.Rows[n].ItemArray.ToIdString(dt)}>");
                            }
                        }
                    }
                }
                else
                {
                    if (string.IsNullOrEmpty(str))
                    {
                        Failed($"[{keyword}] Null이 허용되지 않는데, 비어있는 데이터를 발견했습니다. <Table:{dt.TableName}, 컬럼:{dt.Columns[index].ColumnName} " +
                            $"참조Id:{refId}, Key:{str}, Id:{dt.Rows[n].ItemArray.ToIdString(dt)}>");
                    }

                    if (false == hashSet.Contains(str))
                    {
                        Failed($"[{keyword}] 참조 테이블 내에서 유효한 값이 없습니다. <Table:{dt.TableName}, 컬럼:{dt.Columns[index].ColumnName} " +
                            $"참조Id:{refId}, Key:{str}, Id:{dt.Rows[n].ItemArray.ToIdString(dt)}>");
                    }
                }
            }
            return true;
        }

        public static void ValidationConst(this DataTable dt, string keyword, List<string?> dataTypes)
        {
            var enums = dataTypes.ConvertToReferenceEnum();
            foreach (var e in enums)
            {
                var index = e.Key;
                var hashSet = dt.LoadEnum(keyword, e.Key, e.Value, true);
                if (hashSet == null)
                {
                    Failed($"[{keyword}] HashSet 변환에 실패했습니다. <Table:{dt.TableName}, enum:{e}>");
                    continue;
                }

                var rowDatas = new List<string>();
                for (int n = 0; n < dt.Rows.Count; ++n)
                {
                    var obj = dt.Rows[n].ItemArray.ToArray()[index];
                    if (obj == null)
                        continue;

                    var str = obj.ToStringValue();

                    rowDatas.Add(str);

                    if (false == hashSet.Contains(str))
                    {
                        Failed($"[{keyword}] Enum 참조에 실패했습니다. <Table:{dt.TableName}, enum:{e}, Key:{str}, Id:{dt.Rows[n].ItemArray.ToIdString(dt)}>");
                    }
                }

                var duplicates = rowDatas.GroupBy(x => x)
                    .Where(g => g.Count() > 1)
                    .Select(y => y.Key)
                    .ToList();

                foreach (var col in duplicates)
                {
                    Failed($"컬럼 중복이 발견됐습니다. Const 테이블은 반드시 Enum을 한번만 사용해야 합니다. <Table:{dt.TableName}, Column:{col}>");
                }

                var datas = rowDatas.GroupBy(x => x)
                    .Select(y => y.Key)
                    .ToList();

                if (datas.Count != hashSet.Count)
                {
                    Failed($"Const 테이블은 사용된 Enum의 모든 컬럼을 사용해야 합니다. <Table:{dt.TableName}>");
                }
            }
        }


        public static void ValidateReferenceEnum(this DataTable dt, string keyword, List<string?> dataTypes)
        {
            var enums = dataTypes.ConvertToReferenceEnum();
            foreach (var e in enums)
            {
                var index = e.Key;
                if (e.Value.StartsWith("List"))
                {
                    dt.ValidateReferenceEnum(keyword, index, e.Value.ExtractDataTypeInList(), true);
                }
                else
                {
                    dt.ValidateReferenceEnum(keyword, index, e.Value, false);
                }
            }
        }

        private static HashSet<string>? LoadEnum(this DataTable dt, string keyword, int index, string e, bool checkDuplicate = false)
        {
            var path = string.Format($"output/enumJson/{e.RemoveSpecialCharacters()}.Json");
            if (!File.Exists(path))
            {
                Failed($"[{keyword}] Enum 파일이 존재하지 않습니다. 없는 Enum을 사용했을 가능성이 높습니다. <Table:{dt.TableName}, enum:{e}, Column:{dt.Columns[index].ColumnName}, Path:{path}>");
                return null;
            }

            var json = File.ReadAllText(path);
            var list = JsonConvert.DeserializeObject<List<Dictionary<string, object>>>(json);
            if (list == null)
            {
                Failed($"[{keyword}] JSON을 변환하는 데에 실패했습니다. <Table:{dt.TableName}, enum:{e}, Column:{dt.Columns[index].ColumnName}, Path:{path}>");
                return null;
            }

            return list.ToHashSet(dt.TableName, "Text", checkDuplicate);
        }

        public static bool ValidateReferenceEnum(this DataTable dt, string keyword, int index, string e, bool isList)
        {
            var hashSet = dt.LoadEnum(keyword, index, e);
            if (hashSet == null)
                return false;

            for (int n = 0; n < dt.Rows.Count; ++n)
            {
                var obj = dt.Rows[n].ItemArray.ToArray()[index];
                if (obj == null)
                {
                    Failed($"[{keyword}] 문자열로 변환에 실패했습니다. 참조에 실패했습니다. <Table:{dt.TableName}, enum:{e}>");
                    return false;
                }

                var str = obj.ToStringValue();

                if (e.EndsWith("?") && string.IsNullOrEmpty(str))
                {
                    continue;
                }

                if (isList)
                {
                    if (obj.GetType() != typeof(DBNull))
                    {
                        var strings = JsonConvert.DeserializeObject<List<string>>(str);
                        if (strings == null)
                        {
                            Failed($"[{keyword}] 문자열로 변환에 실패했습니다. 참조에 실패했습니다. <Table:{dt.TableName}, enum:{e}>");
                            return false;
                        }

                        foreach (var checkKey in strings)
                        {
                            if (false == hashSet.Contains(checkKey))
                            {
                                Failed($"[{keyword}] Enum 참조에 실패했습니다. <Table:{dt.TableName}, enum:{e}, Key:{checkKey}, Id:{dt.Rows[n].ItemArray.ToIdString(dt)}>");
                            }
                        }
                    }
                }
                else
                {
                    if (false == hashSet.Contains(str))
                    {
                        Failed($"[{keyword}] Enum 참조에 실패했습니다. <Table:{dt.TableName}, enum:{e}, Key:{str}, Id:{dt.Rows[n].ItemArray.ToIdString(dt)}>");
                    }
                }
            }
            return true;
        }

        private static string ToIdString(this object?[]? objs, DataTable dt)
        {
            var index = dt.Columns.IndexOf("Id");
            if (index == -1)
            {
                return string.Empty;
            }

            if (objs == null)
            {
                return string.Empty;
            }

            var obj = objs[index];
            if (obj == null)
            {
                return string.Empty;
            }

            return obj.ToSafeString();
        }

        public static bool IsPrimitiveType(string qualifiedTypeName)
        {
            switch (qualifiedTypeName)
            {
                case "string":
                case "string?":
                case "float":
                case "float?":
                case "double":
                case "double?":
                case "int":
                case "int?":
                case "DateTime":
                case "DateTime?":
                case "TimeSpan":
                case "TimeSpan?":
                case "bool":
                case "bool?":
                case "Rectangle":
                case "Rectangle?":
                case "DayOfWeek":
                case "DayOfWeek?":
                    return true;
                default:
                    return false;
            }
        }

        public static bool ValidationConvert(string qualifiedTypeName, string str)
        {
            switch (qualifiedTypeName)
            {
                case "string":
                case "string?":
                    return true;
                case "float":
                    return float.TryParse(str, out var _);
                case "float?":
                    return float.TryParse(str, out var _) || string.IsNullOrEmpty(str);
                case "double":
                    return double.TryParse(str, out var _);
                case "double?":
                    return double.TryParse(str, out var _) || string.IsNullOrEmpty(str);
                case "int":
                    return int.TryParse(str, out var _);
                case "int?":
                    return int.TryParse(str, out var _) || string.IsNullOrEmpty(str);
                case "DateTime":
                    return DateTime.TryParse(str, out var _);
                case "DateTime?":
                    return DateTime.TryParse(str, out var _) || string.IsNullOrEmpty(str);
                case "TimeSpan":
                    return TimeSpan.TryParse(str, out var _);
                case "TimeSpan?":
                    return TimeSpan.TryParse(str, out var _) || string.IsNullOrEmpty(str);
                case "bool":
                    return bool.TryParse(str, out var _);
                case "bool?":
                    return bool.TryParse(str, out var _) || string.IsNullOrEmpty(str);
                case "Rectangle":
                    return null != new RectangleConverter().ConvertFromString(str);
                case "Rectangle?":
                    return null != new RectangleConverter().ConvertFromString(str) || string.IsNullOrEmpty(str);
                case "DayOfWeek":
                    return Enum.TryParse(typeof(DayOfWeek), str, out var _);
                case "DayOfWeek?":
                    return Enum.TryParse(typeof(DayOfWeek), str, out var _) || string.IsNullOrEmpty(str);
                default:
                    return false;
            }
        }

        public static bool IsListType(string qualifiedTypeName)
        {
            return qualifiedTypeName.StartsWith("List") && (qualifiedTypeName.EndsWith(">") || qualifiedTypeName.EndsWith(">?"));
        }

        public static bool IsKeyword(string dataType)
        {
            return dataType.StartsWith("$") || dataType.StartsWith("!") || dataType.StartsWith("~") || dataType.StartsWith("%");
        }

        public static void ValidateDataType(this DataTable dt, string keyword, List<string?> dataTypes)
        {
            var dataTypeStrings = dataTypes.ConvertAll(x => x != null ? (string)x : String.Empty);
            for (int index = 0; index < dataTypeStrings.Count; ++index)
            {
                var dataType = dataTypeStrings[index];
                if (IsPrimitiveType(dataType) || IsKeyword(dataType))
                {
                    continue;
                }

                if (IsListType(dataType))
                {
                    if (dataType.Contains("Type") && dataType.Contains("?>"))
                    {
                        Failed($"[{keyword}] List타입이고, Enum을 담고 있다면 Nullable이 아니어야만 합니다 <Table:{dt.TableName}, Column:{dt.Columns[index].ColumnName}, dataType:{dataType}>");
                    }
                    continue;
                }

                if (File.Exists($"output/{keyword}/{dataType}.Json"))
                {
                    continue;
                }

                if (!dataType.EndsWith("Type") && !dataType.EndsWith("Type?"))
                {
                    Failed($"[{keyword}] 참조 컬럼이 아니고, 기본형 타입, List 타입이 아니라면, Type 또는 Type? 으로 끝나야 합니다. <Table:{dt.TableName}, Column:{dt.Columns[index].ColumnName}, dataType:{dataType}>");
                }
            }
        }

        public static void ValidateDuplicate(this DataTable dt)
        {
            var converted = dt.Columns.Cast<DataColumn>().ToList().ConvertAll(x => Regex.Replace(x.ColumnName, @"_[0-9]+$", string.Empty));
            var duplicates = converted.GroupBy(x => x)
                .Where(g => g.Count() > 1)
                .Select(y => y.Key)
                .ToList();

            foreach (var col in duplicates)
            {
                Failed($"컬럼 중복이 발견됐습니다. <Table:{dt.TableName}, Column:{col}>");
            }
        }

        public static void ValidateTableName(string name)
        {
            if (false == CodeDomProvider.CreateProvider("C#").IsValidIdentifier(name) &&
                            !name.StartsWith("!"))
            {
                Failed($"사용 불가능한 테이블명입니다. <Table:{name}>");
            }
        }

        public static void ValidateColumnName(this DataTable dt)
        {
            foreach (var col in dt.Columns.Cast<DataColumn>()
                .Where(x => !CodeDomProvider.CreateProvider("C#").IsValidIdentifier(x.ColumnName)).ToList())
            {
                Failed($"사용 불가능한 컬럼명입니다. <Table:{dt.TableName}, Column:{col}>");
            }

            foreach (var col in dt.Columns.Cast<DataColumn>()
                .Where(x => !Char.IsUpper(x.ColumnName[0]) || (x.ColumnName.Length > 1 && x.ColumnName.Length == x.ColumnName.Where(n => Char.IsUpper(n)).ToList().Count)))
            {
                Failed($"컬럼명은 카멜 케이스만 지원합니다. <Table:{dt.TableName}, Column:{col}>");
            }

            if (!dt.TableName.StartsWith("!") &&
                !dt.TableName.StartsWith("Const") &&
                false == dt.Columns.Cast<DataColumn>().Where(x => x.ColumnName.Contains("Id")).Any())
            {
                Failed($"테이블에 Id 컬럼이 포함되지 않았습니다. <Table:{dt.TableName}>");
            }
        }

        public static void ValidateId(this DataTable dt, string key, bool checkSpace)
        {
            if (dt.TableName.StartsWith("Const"))
            {
                return;
            }

            var columnIndex = dt.Columns.IndexOf(key);
            if (columnIndex == -1)
            {
                Failed($"Id 컬럼을 찾을 수 없었습니다. <Table:{dt.TableName}, Key:{key}>");
                return;
            }

            var converted = dt.Rows.Cast<DataRow>().Select(r => r[columnIndex]).Select(x => x.ToString());
            foreach (var row in converted.GroupBy(x => x)
                .Where(g => g.Count() > 1)
                .Select(y => y.Key).ToList())
            {
                Failed($"중복된 행이 발견되었습니다. <Table:{dt.TableName}, Row:{row}>");
            }

            if (checkSpace)
            {
                foreach (var row in converted.Where(x => x!.Contains(" ")).ToList())
                {
                    Failed($"Id에 공백이 포함되었습니다. <Table:{dt.TableName}, Row:{row}>");
                }
            }

            foreach (var row in dt.Rows.Cast<DataRow>().Select(r => r[columnIndex]).Where(x => x.GetType() == typeof(DBNull)))
            {
                Failed($"Id가 비어있는 행이 발견되었습니다. <Table:{dt.TableName}, Row:{row}>");
            }
        }
    }
}

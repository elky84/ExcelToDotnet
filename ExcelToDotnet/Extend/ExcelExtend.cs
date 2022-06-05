using ExcelDataReader;
using Newtonsoft.Json;
using System.Data;
using System.Numerics;

namespace ExcelToDotnet.Extend

{
    public static class ExcelExtend
    {
        public static KeyValuePair<int, int> FindPosition(this List<List<object?>> list2D, int startX, int startY, string keyword)
        {
            var y = list2D.FindIndex(startY, n => n.Contains(keyword));
            if (y == -1)
            {
                return new KeyValuePair<int, int>(startY, y);
            }

            var x = list2D[y].FindIndex(startX, n => (n != null ? n.ToString() : string.Empty) == keyword);
            return new KeyValuePair<int, int>(x, y);
        }

        public static KeyValuePair<int, int> FindEndField(this List<List<object?>> list2D, int startX, int startY, string keyword)
        {
            var x = list2D[startY].FindIndex(startX, n => (n != null ? n.ToString() : string.Empty) == keyword);
            return new KeyValuePair<int, int>(x, startY);
        }

        public static KeyValuePair<int, int> FindEndTable(this List<List<object?>> list2D, int startX, int startY, string keyword)
        {
            var y = list2D.FindIndex(startY, n =>
            {
                if (n == null || n.Count <= startX)
                    return false;

                var str = n[startX] as string;
                if (str == null)
                {
                    throw new Exception($"not found string. <keyword:{keyword}> <x:{startX}> <y:{startY}> <list:{n}>");
                }
                return str!.ToString() == keyword;
            });
            return new KeyValuePair<int, int>(startX, y);
        }

        public static bool FilterColumn(IExcelDataReader rowReader, int columnIndex)
        {
            if (rowReader[columnIndex] == null)
            {
                return false;
            }

            var str = rowReader[columnIndex].ToString();
            if (str == null)
            {
                return false;
            }

            if (columnIndex == 0) // 0번째 인덱스는 Row의 예외 처리
            {
                return true;
            }

            return !str.Contains('#');
        }

        public static bool FilterColumnSheet(IExcelDataReader rowReader, int columnIndex)
        {
            if (rowReader[columnIndex] == null)
            {
                return false;
            }

            var str = rowReader[columnIndex].ToString();
            if (str == null)
            {
                return false;
            }

            if (columnIndex == 0) // 0번째 인덱스는 Row의 예외 처리
            {
                return true;
            }

            return !str.Contains('#');
        }

        public static bool FilterRow(IExcelDataReader rowReader)
        {
            if (rowReader.FieldCount <= 0 || rowReader[0] == null)
            {
                return false;
            }

            var str = rowReader[0].ToString();
            if (str == null)
            {
                return false;
            }

            return !str.Contains('#');
        }

        public static object? Get(this object?[] objs, int index)
        {
            if (objs == null)
                return null;

            if (objs.Length >= index)
                return null;

            return objs[index];
        }

        public static IEnumerable<DataTable> ParseExcelToDataTable(this Stream document)
        {
            using (var reader = ExcelReaderFactory.CreateReader(document))
            {
                var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    FilterSheet = (tableReader, sheetIndex) => !tableReader.Name.Contains("Enum"),
                    ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                    {
                        UseHeaderRow = true,
                        FilterColumn = (rowReader, columnIndex) => FilterColumn(rowReader, columnIndex),
                        FilterRow = (rowReader) => FilterRow(rowReader)
                    }
                });
                return result.Tables.Cast<DataTable>();
            }
        }


        public static IEnumerable<DataTable> ParseExcelToDataTableEnum(this Stream document)
        {
            using (var reader = ExcelReaderFactory.CreateReader(document))
            {
                var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    FilterSheet = (tableReader, sheetIndex) => tableReader.Name.Contains("Enum")
                });
                return result.Tables.Cast<DataTable>();
            }
        }

        public static Vector3? StringToVector3(this string sVector)
        {
            if (sVector.StartsWith("<") && sVector.EndsWith(">"))
            {
                sVector = sVector[1..^1];
            }
            var sArray = sVector.Split(',');
            if (sArray.Length != 3)
            {
                return null;
            }

            return new Vector3(
                float.Parse(sArray[0]),
                float.Parse(sArray[1]),
                float.Parse(sArray[2]));
        }

        public static Vector2? StringToVector2(this string sVector)
        {
            if (sVector.StartsWith("<") && sVector.EndsWith(">"))
            {
                sVector = sVector[1..^1];
            }
            var sArray = sVector.Split(',');
            if (sArray.Length != 2)
            {
                return null;
            }

            return new Vector2(
                float.Parse(sArray[0]),
                float.Parse(sArray[1]));
        }

        public static IEnumerable<Dictionary<string, object>> MapDatasetData(this DataTable dt, List<string>? dataTypes = null)
        {
            foreach (DataRow dr in dt.Rows)
            {
                var row = new Dictionary<string, object>();
                foreach (DataColumn col in dt.Columns)
                {
                    if (int.TryParse(dr[col].ToString(), out var value))
                    {
                        row.Add(col.ColumnName, value);
                    }
                    else
                    {
                        var dataType = dataTypes != null ? dataTypes[dt.Columns.IndexOf(col)] : string.Empty;
                        if (dataType.StartsWith("List"))
                        {
                            try
                            {
                                var datas = dr[col];
                                if (datas.GetType() != typeof(string) || string.IsNullOrEmpty((string)dr[col]))
                                {
                                    row.Add(col.ColumnName, new List<object?> { });
                                }
                                else
                                {
                                    var objects = JsonConvert.DeserializeObject<List<object?>>((string)dr[col]);
                                    if (objects == null)
                                    {
                                        LogExtend.Color(ConsoleColor.Red, ConsoleColor.Black, $"List 형태의 데이터가 아닙니다. 대괄호[]로 묶여있는지 확인해주세요. " +
                                            $"<Table:{dt.TableName}, DataType: {dataType}, Column:{dr[col]}>");
                                    }
                                    else
                                    {
                                        row.Add(col.ColumnName, objects);
                                    }
                                }
                            }
                            catch (Exception e)
                            {
                                LogExtend.Color(ConsoleColor.Red, ConsoleColor.Black, $"List 형태의 데이터가 아닙니다. 대괄호[]로 묶여있는지, 문자열의 경우 따옴표로 묶여있는지, 콤마 구분자를 잘 사용했는지 등을 확인해주세요. " +
                                    $"<Table:{dt.TableName}, DataType: {dataType}, Column:{dr[col]}, Exception:{e}>");
                                throw;
                            }
                        }
                        else if (dataType.RemoveSpecialCharacters().StartsWith("Vector2"))
                        {
                            var vector2 = ((string)dr[col]).StringToVector2();
                            if (vector2 == null)
                            {
                                LogExtend.Color(ConsoleColor.Red, ConsoleColor.Black, $"Vector타입은 <> 로 묶여 있어야 합니다. 콤마 구분자도 확인해주세요 " +
                                    $"<Table:{dt.TableName}, DataType: {dataType}, Column:{dr[col]}>");
                                continue;
                            }

                            row.Add(col.ColumnName, vector2);
                        }
                        else if (dataType.RemoveSpecialCharacters().StartsWith("Vector3"))
                        {
                            var vector3 = ((string)dr[col]).StringToVector3();
                            if (vector3 == null)
                            {
                                LogExtend.Color(ConsoleColor.Red, ConsoleColor.Black, $"Vector타입은 <> 로 묶여 있어야 합니다. 콤마 구분자도 확인해주세요 " +
                                    $"<Table:{dt.TableName}, DataType: {dataType}, Column:{dr[col]}>");
                                continue;
                            }

                            row.Add(col.ColumnName, vector3);
                        }
                        else
                        {
                            row.Add(col.ColumnName, dr[col]);
                        }
                    }
                }
                yield return row;
            }
        }

        public static void RemoveCommentRow(this DataTable dt)
        {
            for (int n = 0; n < dt.Rows.Count; ++n)
            {
                var firstColumn = dt.Rows[n].ItemArray.FirstOrDefault();
                if (firstColumn != null &&
                    firstColumn.ToStringValue().StartsWith("#"))
                {
                    dt.Rows.RemoveAt(n);
                }
            }
        }


        public static void RemoveByTag(this DataTable dt, Options opts)
        {
            if (!string.IsNullOrEmpty(opts.RowEndTag))
            {
                for (int n = 0; n < dt.Rows.Count; ++n)
                {
                    if (dt.Rows[n].ItemArray.ToList().Find(c => c != null && c.GetType() == typeof(string) && opts.RowEndTag == (string)c) != null)
                    {
                        var rowCount = dt.Rows.Count;
                        for (int x = n; x < rowCount; ++x)
                        {
                            dt.Rows.RemoveAt(n);
                        }
                    }
                }
            }

            if (!string.IsNullOrEmpty(opts.ColumnEndTag))
            {
                for (int n = 0; n < dt.Columns.Count; ++n)
                {
                    if (dt.Columns[n].ColumnName == opts.ColumnEndTag)
                    {
                        var colCount = dt.Columns.Count;
                        for (int x = n; x < colCount; ++x)
                        {
                            dt.Columns.RemoveAt(n);
                        }
                    }
                }
            }
        }

        public static DataTable ToDataEnumTable(this List<List<string>> list)
        {
            DataTable tmp = new DataTable();
            tmp.Columns.Add("Code", typeof(string));
            tmp.Columns.Add("Text", typeof(string));

            foreach (List<string> row in list)
            {
                tmp.Rows.Add(row.ToArray());
            }
            return tmp;
        }

    }
}

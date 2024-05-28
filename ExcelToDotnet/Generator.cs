using ExcelToDotnet.Extend;
using Newtonsoft.Json;
using System.Data;
using System.Text;

namespace ExcelToDotnet
{
    public static class Generator
    {
        public static void Execute(Options opts)
        {
            PrepareOutputDirectories(opts);

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            List<string> files = GetInputFiles(opts);

            if (!opts.Validation && files.Count == 0)
            {
                LogExtend.Color(ConsoleColor.Yellow, ConsoleColor.Magenta, $"Not found input files. <InputDirectory: {opts.InputDirectory}, Files: {string.Join(",", opts.InputFiles)}>");
                Environment.Exit(ErrorCode.NO_INPUT_FILES);
            }

            foreach (var fileName in files)
            {
                ProcessFile(fileName, opts);
            }
        }

        private static void PrepareOutputDirectories(Options opts)
        {
            if ((opts.CleanUp || opts.Wide) && Directory.Exists(opts.Output))
            {
                opts.Output.DeleteDirectory();
            }

            Directory.CreateDirectory(opts.Output);
            Directory.CreateDirectory(Path.Combine(opts.Output, "enum"));
            Directory.CreateDirectory(Path.Combine(opts.Output, "enumJson"));
        }

        private static List<string> GetInputFiles(Options opts)
        {
            if (!string.IsNullOrEmpty(opts.InputDirectory))
            {
                return Directory.GetFiles(opts.InputDirectory)
                    .Where(x => x.ToLower().EndsWith(".xlsx") && !Path.GetFileName(x).StartsWith("~$"))
                    .ToList();
            }
            return opts.InputFiles.ToList();
        }

        private static void ProcessFile(string fileName, Options opts)
        {
            if (opts.Enum || opts.Wide)
            {
                GenerateEnum(fileName, opts);
            }

            if (!opts.Enum || opts.Wide)
            {
                GenerateTable(fileName, opts);
            }
        }

        public static void GenerateTable(string fileName, Options opts)
        {
            var tempFileName = CreateTemporaryFile(fileName);
            try
            {
                using var stream = File.Open(tempFileName, FileMode.Open, FileAccess.Read);
                var dataTables = stream.ParseExcelToDataTable();
                foreach (var (tableName, dataTable) in dataTables.ToDictionary(x => x.TableName))
                {
                    if (!tableName.Contains('#'))
                    {
                        GenerateTableFromData(fileName, opts, tableName.Replace("!", string.Empty), dataTable);
                    }
                }
            }
            catch (IOException exception)
            {
                HandleIOException(fileName, exception, tempFileName);
            }
            finally
            {
                File.Delete(tempFileName);
            }
        }

        public static void GenerateEnum(string fileName, Options opts)
        {
            var tempFileName = CreateTemporaryFile(fileName);
            try
            {
                using var stream = File.Open(tempFileName, FileMode.Open, FileAccess.Read);
                var dataTables = stream.ParseExcelToDataTableEnum();
                foreach (var (_, dataTable) in dataTables.ToDictionary(x => x.TableName))
                {
                    GenerateEnumFromData(opts, dataTable);
                }
            }
            catch (IOException exception)
            {
                HandleIOException(fileName, exception, tempFileName);
            }
            finally
            {
                File.Delete(tempFileName);
            }
        }

        private static string CreateTemporaryFile(string fileName)
        {
            var tempFileName = $"{fileName}.temp";
            File.Copy(fileName, tempFileName, true);
            return tempFileName;
        }

        private static void HandleIOException(string fileName, IOException exception, string tempFileName)
        {
            LogExtend.Color(ConsoleColor.Yellow, ConsoleColor.Magenta, $"IO Exception. <File: {fileName}, Reason: {exception.Message}>");
            File.Delete(tempFileName);
            Environment.Exit(ErrorCode.OPENED_EXCEL);
        }

        private static void GenerateTableFromData(string fileName, Options opts, string tableName, DataTable dataTable)
        {
            if (dataTable.Rows.Count < 2)
            {
                Validation.Failed($"스키마 정보가 유효하지 않습니다. 상위 3개의 행은 필수로 존재해야 합니다. !! 1행: 컬럼 이름, 2행: 데이터 타입, 3행: 타겟 (다중 타겟시 |로 구분자) !! <file:{fileName}, table:{tableName} row count: {dataTable.Rows.Count}>");
                return;
            }

            dataTable.RemoveByTag(opts);

            ValidateDataTable(dataTable);

            var dataTypes = dataTable.Rows[0].ItemArray.Select(x => (x as string)!).ToList();
            var targetsArray = dataTable.Rows[1].ItemArray.Select(x => x?.ToString() ?? "").ToArray();

            RemoveHeaderRows(dataTable);

            var dataTableEx = SplitDataTableByTarget(dataTable, dataTypes, targetsArray, opts);

            ValidateAndGenerateOutput(fileName, tableName, opts, dataTableEx);
        }

        private static void ValidateDataTable(DataTable dataTable)
        {
            Validation.ValidateTableName(dataTable.TableName);
            dataTable.ValidateDuplicate();
            dataTable.ValidateColumnName();
        }

        private static void RemoveHeaderRows(DataTable dataTable)
        {
            dataTable.Rows.RemoveAt(0);
            dataTable.Rows.RemoveAt(0);
        }

        private static Dictionary<string, DataTableEx> SplitDataTableByTarget(DataTable dataTable, List<string> dataTypes, string[] targetsArray, Options opts)
        {
            if (targetsArray == null)
                throw new NullReferenceException($"targetsArray is null");

            var targetGroups = targetsArray.GroupBy(x => x).ToList();
            var dataTableEx = new Dictionary<string, DataTableEx>();

            foreach (var targetObj in targetGroups)
            {
                var targets = targetObj.Key.Split("|");
                foreach (var target in targets)
                {
                    if (!dataTableEx.ContainsKey(target))
                    {
                        dataTableEx[target] = new DataTableEx { Target = target, DataTable = dataTable.Copy(), DataTypes = dataTypes.ToList() };
                        Directory.CreateDirectory(Path.Combine(opts.Output, target));
                    }
                }
            }

            RemoveUnusedColumns(dataTableEx, targetsArray);

            return dataTableEx;
        }

        private static void RemoveUnusedColumns(Dictionary<string, DataTableEx> dataTableEx, object?[] targetsArray)
        {
            for (int n = 0; n < targetsArray.Length; ++n)
            {
                var target = targetsArray[n]?.ToString() ?? "";
                var targets = target.Split("|");

                foreach (var dataTable in dataTableEx.Where(x => !targets.Contains(x.Key)))
                {
                    var deleteAt = dataTable.Value.Table.Columns.Count - targetsArray.Length + n;
                    dataTable.Value.Table.Columns.RemoveAt(deleteAt);
                    dataTable.Value.Types.RemoveAt(deleteAt);
                }
            }
        }

        private static void ValidateAndGenerateOutput(string fileName, string tableName, Options opts, Dictionary<string, DataTableEx> dataTableEx)
        {
            if (!opts.Validation || opts.Wide)
            {
                foreach (var dataTable in dataTableEx)
                {
                    GenerateCsAndJson(fileName, dataTable.Value.Target, tableName, opts, dataTable.Value.Table, dataTable.Value.Types);
                }
            }

            if (opts.Validation)
            {
                foreach (var dataTable in dataTableEx)
                {
                    ValidateDataTableContent(tableName, dataTable.Value);
                }
            }
        }

        private static void ValidateDataTableContent(string tableName, DataTableEx dataTableEx)
        {
            var dt = dataTableEx.Table;
            var target = dataTableEx.Target;
            var dataTypes = dataTableEx.Types;

            dt.ValidateValue(target, dataTypes);

            if (!tableName.StartsWith("Const"))
            {
                dt.ValidateReference(target, dataTypes);
                dt.ValidateReferenceEnum(target, dataTypes);
                dt.ValidateSubIndex(target, dataTypes);
                dt.ValidationProbability(target, dataTypes);
            }
            else
            {
                dt.ValidationConst(target, dataTypes);
            }
        }

        private static void GenerateEnumFromData(Options opts, DataTable dataTable)
        {
            var list2D = dataTable.AsEnumerable().Select(row => row.ItemArray.ToList()).ToList();
            GenerateEnum(new KeyValuePair<int, int>(0, 0), opts, list2D);
        }

        private static bool GenerateEnum(KeyValuePair<int, int> startPos, Options opts, List<List<object?>> list2D)
        {
            if (list2D == null)
                return false;

            var start = list2D.FindPosition(startPos.Key, startPos.Value, opts.BeginTag);
            if (start.Value == -1)
            {
                return true;
            }

            if (start.Key == -1)
            {
                return GenerateEnum(new KeyValuePair<int, int>(0, start.Value + 1), opts, list2D);
            }

            var endTable = list2D.FindEndTable(start.Key, start.Value, opts.RowEndTag);
            var endField = list2D.FindEndField(start.Key, start.Value, opts.ColumnEndTag);

            var list = ExtractEnumData(list2D, start, endTable, endField);

            var tableName = list2D[start.Value][start.Key + 1];
            if (tableName == null)
            {
                Validation.Failed($"테이블 명을 가져오는 데에 실패했습니다. <X:{start.Value},Y:{start.Key + 1}>");
                return false;
            }

            GenerateEnum(tableName.ToSafeString(), opts, list.ToDataEnumTable());
            return GenerateEnum(new KeyValuePair<int, int>(start.Key + 1, start.Value), opts, list2D);
        }

        private static List<List<string>> ExtractEnumData(List<List<object?>> list2D, KeyValuePair<int, int> start, KeyValuePair<int, int> endTable, KeyValuePair<int, int> endField)
        {
            var list = new List<List<string>>();
            for (int y = start.Value + 1; y < endTable.Value; ++y)
            {
                var innerList = new List<string>();
                for (int x = start.Key; x < endField.Key; ++x)
                {
                    innerList.Add(list2D[y][x]?.ToSafeString() ?? string.Empty);
                }
                list.Add(innerList);
            }
            return list;
        }

        public static bool GenerateEnum(string tableName, Options opts, DataTable dt)
        {
            Validation.ValidateTableName(tableName);

            dt.TableName = tableName;

            dt.ValidateId("Text", false);
            dt.ValidateId("Code", true);

            string jsonFileName = Path.Combine(opts.Output, "enumJson", $"{tableName}.json");
            if (File.Exists(jsonFileName))
            {
                Validation.Failed($"이미 존재하는 Enum을 사용하셨습니다. <FileName:{jsonFileName}> <Table:{tableName}>");
                return false;
            }

            WriteJsonFile(jsonFileName, dt.MapDatasetData());

            var patterns = GetPatterns(opts, tableName);

            WriteEnumFile(opts, tableName, dt, patterns);

            return true;
        }

        private static void WriteJsonFile(string jsonFileName, object data)
        {
            using var outputFileJson = new StreamWriter(jsonFileName) { NewLine = "\n" };
            string json = JsonConvert.SerializeObject(data, Formatting.Indented).Replace("\r\n", "\n");
            outputFileJson.Write(json);
        }

        private static Dictionary<string, string> GetPatterns(Options opts, string tableName)
        {
            return new Dictionary<string, string>()
            {
                { "##NAMESPACE##", opts.NameSpace },
                { "##CLASS##", tableName },
                { "##USING##", string.Join(Environment.NewLine, opts.Usings) },
                { "##ATTRIBUTE##", opts.Attribute }
            };
        }

        private static void WriteEnumFile(Options opts, string tableName, DataTable dt, Dictionary<string, string> patterns)
        {
            using (FileStream vStream = File.Create($"{opts.Output}/enum/{tableName}.cs"))
            {
                using StreamWriter outputFile = new StreamWriter(vStream, new UTF8Encoding(true));

                CodeTemplate.Enum().GenerateEnumCode(outputFile, dt, patterns);
            }
        }

        public static void GenerateCsAndJson(string fileName, string keyword, string tableName, Options opts, DataTable dt, List<string> dataTypes)
        {
            if (dataTypes.Count == 0) return;

            string jsonFileName = Path.Combine(opts.Output, keyword, $"{tableName}.json");
            if (File.Exists(jsonFileName))
            {
                Validation.Failed($"이미 존재하는 테이블 명을 사용하셨습니다. <ExcelFileName{fileName} jsonFileName:{jsonFileName}, Table:{tableName}>");
                return;
            }

            WriteJsonFile(jsonFileName, dt.MapDatasetData(dataTypes.ConvertAll(x => x.ToString() ?? "")));

            var patterns = GetPatterns(opts, tableName, keyword);

            WriteCsFile(opts, keyword, tableName, dt, dataTypes, patterns);
        }

        private static Dictionary<string, string> GetPatterns(Options opts, string tableName, string keyword)
        {
            return new Dictionary<string, string>()
            {
                { "##NAMESPACE##", $"{opts.NameSpace}.{keyword}" },
                { "##CLASS##", tableName },
                { "##USING##", string.Join(Environment.NewLine, opts.Usings) }
            };
        }

        private static void WriteCsFile(Options opts, string keyword, string tableName, DataTable dt, List<string> dataTypes, Dictionary<string, string> patterns)
        {
            string filePath = Path.Combine(opts.Output, keyword, $"{tableName}.cs");
            using var outputFile = new StreamWriter(filePath);
            CodeTemplate.Class().GenerateCode(outputFile, dt, dataTypes, patterns, opts.Nullable);
        }
    }
}

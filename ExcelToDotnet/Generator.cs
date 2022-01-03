using ExcelToDotnet.Extend;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToDotnet
{
    public static class Generator
    {

        public static void Execute(Options opts)
        {
            if (opts.CleanUp && Directory.Exists(opts.Output))
            {
                opts.Output.DeleteDirectory();
            }

            Directory.CreateDirectory(opts.Output);
            Directory.CreateDirectory(opts.Output + "/enum");
            Directory.CreateDirectory(opts.Output + "/enumJson");

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            List<string> files = !string.IsNullOrEmpty(opts.InputDirectory) ?
                Directory.GetFiles(opts.InputDirectory).Where(x => x.ToLower().EndsWith(".xlsx") && !Path.GetFileName(x).StartsWith("~$")).ToList() :
                opts.InputFiles.ToList();

            if (files.Count <= 0)
            {
                LogExtend.Color(ConsoleColor.Yellow, ConsoleColor.Magenta, $"Not found input files. <InputDirectory: {opts.InputDirectory}, Files: {string.Join(",", opts.InputFiles)}>");
                Environment.Exit(ErrorCode.NO_INPUT_FILES);
            }

            foreach (var fileName in files)
            {
                if (opts.Enum)
                {
                    Generator.GenerateEnum(fileName, opts);
                }
                else
                {
                    Generator.GenerateTable(fileName, opts);
                }
            }
        }


        public static void GenerateTable(string fileName, Options opts)
        {
            var tempFileName = fileName + ".temp";
            try
            {
                File.Copy(fileName, tempFileName, true);

                using var stream = File.Open(tempFileName, FileMode.Open, FileAccess.Read);
                var dataTables = stream.ParseExcelToDataTable();
                foreach (var val in dataTables.ToDictionary(x => x.TableName))
                {
                    var tableName = val.Key;
                    if (tableName.Contains('#'))
                    {
                        continue;
                    }

                    GenerateTable(fileName, opts, tableName, val.Value);
                }

                File.Delete(tempFileName);
            }
            catch (IOException exception)
            {
                LogExtend.Color(ConsoleColor.Yellow, ConsoleColor.Magenta, $"IO Exception. <File: {fileName}, Reason: {exception.Message}>");
                File.Delete(tempFileName);
                Environment.Exit(ErrorCode.OPENED_EXCEL);
            }
        }

        public static void GenerateEnum(string fileName, Options opts)
        {
            var tempFileName = fileName + ".temp";
            try
            {
                File.Copy(fileName, tempFileName, true);

                using var stream = File.Open(tempFileName, FileMode.Open, FileAccess.Read);
                var dataTables = stream.ParseExcelToDataTableEnum();
                foreach (var val in dataTables.ToDictionary(x => x.TableName))
                {
                    GenerateEnum(opts, val.Value);
                }
                File.Delete(tempFileName);
            }
            catch (IOException exception)
            {
                LogExtend.Color(ConsoleColor.Yellow, ConsoleColor.Magenta, $"IO Exception. <File: {fileName}, Reason: {exception.Message}>");
                File.Delete(tempFileName);
                Environment.Exit(ErrorCode.OPENED_EXCEL);
            }
        }

        private static void GenerateEnum(Options opts, DataTable dt)
        {
            GenerateEnum(new KeyValuePair<int, int>(0, 0), opts, dt.AsEnumerable().Select(row => row.ItemArray.ToList()).ToList());
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

            var list = new List<List<string>>();
            for (int y = start.Value + 1; y < endTable.Value; ++y)
            {
                var innerList = new List<string>();
                for (int x = start.Key; x < endField.Key; ++x)
                {
                    if (list2D[y][x] == null)
                    {
                        innerList.Add(String.Empty);
                    }
                    else
                    {
                        innerList.Add(list2D[y][x].ToSafeString());
                    }
                }
                list.Add(innerList);
            }

            var tableName = list2D[start.Value][start.Key + 1];
            if (tableName == null)
            {
                Validation.Failed($"테이블 명을 가져오는 데에 실패했습니다. <X:{start.Value},Y:{start.Key + 1}>");
                return false;
            }

            GenerateEnum(tableName.ToSafeString(), opts, list.ToDataEnumTable());
            return GenerateEnum(new KeyValuePair<int, int>(start.Key + 1, start.Value), opts, list2D);
        }


        public static bool GenerateEnum(string tableName, Options opts, DataTable dt)
        {
            if (!opts.Ignore)
            {
                Validation.ValidateTableName(tableName);
            }

            dt.TableName = tableName;

            if (!opts.Ignore)
            {
                dt.ValidateId("Text", false);
                dt.ValidateId("Code", true);
            }

            string jsonFileName = $"{opts.Output}/enumJson/{tableName}.json";
            if (File.Exists(jsonFileName))
            {
                Validation.Failed($"이미 존재하는 Enum을 사용하셨습니다. <FileName:{jsonFileName}> <Table:{tableName}>");
                return false;
            }

            using StreamWriter outputFileJson = new StreamWriter(jsonFileName) { NewLine = "\n" };

            string json = JsonConvert.SerializeObject(dt.MapDatasetData(), Formatting.Indented);
            json = json.Replace("\r\n", "\n");
            outputFileJson.Write(json);

            var patterns = new Dictionary<string, string>() {
                                { "##NAMESPACE##", opts.NameSpace},
                                { "##CLASS##", tableName},
                                { "##USING##", string.Join( Environment.NewLine, opts.Usings) },
                                { "##ATTRIBUTE##", opts.Attribute }
                            };

            using (FileStream vStream = File.Create($"{opts.Output}/enum/{tableName}.cs"))
            {
                using StreamWriter outputFile = new StreamWriter(vStream, new UTF8Encoding(true));

                CodeTemplate.Enum().GenerateEnumCode(outputFile, dt, patterns);
            }
            return true;
        }

        public static void GenerateTable(string fileName, Options opts, string tableName, DataTable dt)
        {
            if (dt.Rows.Count < 2)
            {
                Validation.Failed($"스키마 정보가 유효하지 않습니다. 상위 3개의 행은 필수로 존재해야 합니다. !! 1행: 컬럼 이름, 2행: 데이터 타입, 3행: 타겟 (다중 타겟시 |로 구분자) !! <file:{fileName}, table:{tableName} row count: {dt.Rows.Count}>");
                return;
            }

            dt.RemoveByTag(opts);

            if (!opts.Ignore)
            {
                Validation.ValidateTableName(dt.TableName);
                dt.ValidateDuplicate();
                dt.ValidateColumnName();
            }

            var dataTypes = dt.Rows[0].ItemArray.Select(x => (x as string)!).ToList();
            var targetsArray = dt.Rows[1].ItemArray;

            // 1행은 컬럼으로 빠졌고, 2행, 3행을 지워야되는데, 0번으로 지우면 앞으로 당겨져서 0번으로 두번 지움
            dt.Rows.RemoveAt(0);
            dt.Rows.RemoveAt(0);

            if (!opts.Ignore)
            {
                dt.ValidateId("Id", false);
            }

            var targetGroups = targetsArray.ToList().GroupBy(x => x).Select(x => x.Key).ToList();
            var dataTableEx = new Dictionary<string, DataTableEx>();

            foreach (var targetObj in targetGroups)
            {
                var targets = (targetObj as string)!.Split("|");
                foreach (var target in targets)
                {
                    if (!dataTableEx.ContainsKey(target!.ToString()))
                    {
                        dataTableEx.Add(target!.ToString(), new DataTableEx { Target = target, DataTable = dt.Copy(), DataTypes = dataTypes.ToList() });
                        Directory.CreateDirectory(opts.Output + "/" + target);
                    }
                }
            }

            for (int n = 0; n < targetsArray.Length; ++n)
            {
                var target = targetsArray[n] as string;
                var targets = target!.Split("|");

                foreach (var dataTable in dataTableEx.Where(x => !targets.Contains(x.Key)))
                {
                    var deleteAt = dataTable.Value!.DataTable!.Columns.Count - targetsArray.Length + n;
                    dataTable.Value!.DataTable!.Columns.RemoveAt(deleteAt);
                    dataTable.Value!.DataTypes!.RemoveAt(deleteAt);
                }
            }

            if (!opts.Ignore)
            {
                foreach (var dataTable in dataTableEx)
                {
                    dataTable.Value!.DataTable!.ValidateDataType(dataTable.Key, dataTable.Value!.DataTypes!);
                }
            }


            if (!opts.Validation)
            {
                foreach (var dataTable in dataTableEx)
                {
                    GenerateCsAndJson(fileName, dataTable.Value.Target.ToCamelCase(), tableName, opts, dataTable.Value!.DataTable!, dataTable.Value!.DataTypes!);
                }
            }
            else
            {
                if (!opts.Ignore)
                {
                    foreach (var dataTable in dataTableEx)
                    {
                        dataTable.Value!.DataTable!.ValidateValue(dataTable.Value.Target.ToCamelCase(), dataTable.Value!.DataTypes!);

                        if (!tableName.StartsWith("Const"))
                        {
                            dataTable.Value!.DataTable!.ValidateReference(dataTable.Value.Target.ToCamelCase(), dataTable.Value!.DataTypes!);

                            dataTable.Value!.DataTable!.ValidateReferenceEnum(dataTable.Value.Target.ToCamelCase(), dataTable.Value!.DataTypes!);

                            dataTable.Value!.DataTable!.ValidateSubIndex(dataTable.Value.Target.ToCamelCase(), dataTable.Value!.DataTypes!);

                            dataTable.Value!.DataTable!.ValidationProbability(dataTable.Value.Target.ToCamelCase(), dataTable.Value!.DataTypes!);
                        }
                        else
                        {
                            dataTable.Value!.DataTable!.ValidationConst(dataTable.Value.Target.ToCamelCase(), dataTable.Value!.DataTypes!);
                        }
                    }
                }
            }
        }

        public static bool GenerateCsAndJson(string fileName, string keyword, string tableName, Options opts, DataTable dt, List<string?> dataTypes)
        {
            if (dataTypes.Count <= 0)
            {
                return false;
            }

            string jsonFileName = $"{opts.Output}/{keyword}/{tableName}.json";
            if (File.Exists(jsonFileName))
            {
                Validation.Failed($"이미 존재하는 테이블 명을 사용하셨습니다. <ExcelFileName{fileName} jsonFileName:{jsonFileName}, Table:{tableName}>");
                return false;
            }

            using StreamWriter outputFileJson = new StreamWriter(jsonFileName) { NewLine = "\n" };
            string json = JsonConvert.SerializeObject(dt.MapDatasetData(dataTypes.ConvertAll(x => x.ToStringValue()).ToList()), Formatting.Indented);
            json = json.Replace("\r\n", "\n");
            outputFileJson.Write(json);

            var patterns = new Dictionary<string, string>() {
                                { "##NAMESPACE##", opts.NameSpace + "." + keyword},
                                { "##CLASS##", tableName},
                                { "##USING##", string.Join( Environment.NewLine, opts.Usings) }
                            };

            using var outputFile = new StreamWriter($"{opts.Output }/{keyword}/{tableName}.cs");

            CodeTemplate.Class().GenerateCode(outputFile, dt, dataTypes, patterns);
            return true;
        }
    }
}

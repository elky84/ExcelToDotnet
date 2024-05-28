using System.Data;

namespace ExcelToDotnet
{
    public class DataTableEx
    {
        public string Target { get; set; } = "";

        public DataTable? DataTable { private get; set; }

        public DataTable Table => DataTable!;

        public List<string>? DataTypes { private get; set; }

        public List<string> Types => DataTypes!;

    }
}

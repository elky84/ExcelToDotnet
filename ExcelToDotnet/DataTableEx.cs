using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToDotnet
{
    public class DataTableEx
    {
        public string Target { get; set; } = "";

        public DataTable? DataTable { get; set; }

        public List<string>? DataTypes { get; set; }
    }
}

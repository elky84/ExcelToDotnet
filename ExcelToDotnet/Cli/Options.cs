using CommandLine;
using System;
using System.Collections.Generic;

namespace ExcelToDotnet.Cli
{
    public partial class Options
    {
        [Option('d', "directory", Required = false, HelpText = "excel file directory.")]
        public string InputDirectory { get; set; } = "";

        [Option('f', "files", Required = false, HelpText = "excel file names.")]
        public IEnumerable<string> InputFiles { get; set; } = new List<string>();

        [Option('n', "namespace", Required = false, HelpText = "namespace")]
        public string NameSpace { get; set; } = "DEFAULT_NAMESPACE";

        [Option('o', "output", Required = false, HelpText = "output directory")]
        public string Output { get; set; } = "output";

        [Option('u', "using", Required = false, HelpText = "using statements")]
        public IEnumerable<string> Usings { get; set; } = new List<string>();

        [Option('c', "cleanup", Required = false, HelpText = "cleanup output directory")]
        public bool CleanUp { get; set; }

        [Option('a', "attribute", Required = false, HelpText = "class attribute")]
        public string Attribute { get; set; } = "";

        [Option('b', "begin tag", Required = false, HelpText = "begin tag")]
        public string BeginTag { get; set; } = ":Begin";

        [Option('r', "rowendtag", Required = false, HelpText = "row end tag")]
        public string RowEndTag { get; set; } = ":End";

        [Option('z', "colendtag", Required = false, HelpText = "column end tag")]
        public string ColumnEndTag { get; set; } = ":End";

        [Option('e', "enum", Required = false, HelpText = "process enum")]
        public bool Enum { get; set; }

        [Option('v', "validation", Required = false, HelpText = "validation mode")]
        public bool Validation { get; set; }

        [Option('l', "nullable", Required = false, HelpText = "set nullable")]
        public bool Nullable { get; set; }

        [Option('w', "wide", Required = false, HelpText = "wide mode (This option runs cleanup, enum, code and json functions at once.)")]
        public bool Wide { get; set; }

        public static void HandleParseError(IEnumerable<Error> errs)
        {
            foreach (var err in errs)
            {
                Console.WriteLine(err.ToString());
            }
        }
    }
}

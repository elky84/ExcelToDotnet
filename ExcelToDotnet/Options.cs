using CommandLine;
using System;
using System.Collections.Generic;

namespace ExcelToDotnet
{
    public partial class Options
    {
        [Option('d', "directory", Required = false, HelpText = "Input excel file directory.")]
        public string InputDirectory { get; set; } = "";

        [Option('f', "files", Required = false, HelpText = "Input excel file names.")]
        public IEnumerable<string> InputFiles { get; set; } = new List<string>();

        [Option('n', "namespace", Required = false, HelpText = "Input namespace")]
        public string NameSpace { get; set; } = "DEFAULT_NAMESPACE";

        [Option('o', "output", Required = false, HelpText = "Input output directory")]
        public string Output { get; set; } = "output";

        [Option('u', "using", Required = false, HelpText = "Input using statements")]
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

        [Option('i', "ignore", Required = false, HelpText = "Ignore general validation")]
        public bool Ignore { get; set; }

        public static void HandleParseError(IEnumerable<Error> errs)
        {
            foreach (var err in errs)
            {
                Console.WriteLine(err.ToString());
            }
        }
    }
}

namespace ExcelToDotnet
{
    public static class CodeTemplate
    {
        public static List<string> Enum()
        {
            return new List<string> {
                        "##USING##",
                        "",
                        "namespace ##NAMESPACE##",
                        "{",
                        "\t##ATTRIBUTE##",
                        "\tpublic enum ##CLASS##",
                        "\t{",
                        "\t##INSERT##",
                        "\t}",
                        "}" };
        }

        public static List<string> Class()
        {
            return new List<string> {
                        "using System;",
                        "using System.Collections.Generic;",
                        "using System.Drawing;",
                        "using Newtonsoft.Json;",
                        "##USING##",
                        "",
                        "namespace ##NAMESPACE##",
                        "{",
                        "\tpublic class ##CLASS##",
                        "\t{",
                        "\t##INSERT##",
                        "\t}",
                        "}" };
        }
    }
}

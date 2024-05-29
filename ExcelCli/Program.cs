using CommandLine;
using ExcelToDotnet.Actor;
using ExcelToDotnet.Cli;
using ExcelToDotnet.Code;
using ExcelToDotnet.Extend;
using System;

namespace ExcelTool
{
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                Parser.Default.ParseArguments<Options>(args)
                    .WithParsed<Options>(opts => Generator.Execute(opts))
                    .WithNotParsed<Options>((errs) => Options.HandleParseError(errs));
            }
            catch (Exception exception)
            {
                LogExtend.Color(ConsoleColor.Yellow, ConsoleColor.Magenta, $"Unhandled exception. <Reason:{exception.Message}> <StackTrace:{exception.StackTrace}>");
                Environment.Exit(ErrorCode.EXCEPTION);
            }
        }
    }
}

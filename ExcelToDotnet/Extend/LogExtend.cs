namespace ExcelToDotnet.Extend
{
    public static class LogExtend
    {
        public static void Color(ConsoleColor bgColor, ConsoleColor fgColor, string format, params object[] args)
        {
            Console.BackgroundColor = bgColor;
            Console.ForegroundColor = fgColor;

            Console.WriteLine(format, args);

            Console.ResetColor();
        }
    }
}

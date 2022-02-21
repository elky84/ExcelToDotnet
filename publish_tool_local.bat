dotnet pack ExcelCli -c Release -o ../DotnetPack

dotnet tool uninstall -g ExcelCli

dotnet tool install -g ExcelCli --add-source ../DotnetPack

pause


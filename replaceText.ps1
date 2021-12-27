$PrevVersion='1.0.3'
$NextVersion='1.0.4'

((Get-Content -path ExcelToDotnet\ExcelToDotnet.csproj -Raw) -replace $PrevVersion,$NextVersion) | Set-Content -Path ExcelToDotnet\ExcelToDotnet.csproj

((Get-Content -path publish_to_github.bat -Raw) -replace $PrevVersion,$NextVersion) | Set-Content -Path publish_to_github.bat

((Get-Content -path publish_to_nuget.bat -Raw) -replace $PrevVersion,$NextVersion) | Set-Content -Path publish_to_nuget.bat

pause
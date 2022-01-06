[![Website](https://img.shields.io/website-up-down-green-red/http/shields.io.svg?label=elky-essay)](https://elky84.github.io)
![Made with](https://img.shields.io/badge/made%20with-.NET6-blue.svg)

![GitHub forks](https://img.shields.io/github/forks/elky84/ExcelToDotnet.svg?style=social&label=Fork)
![GitHub stars](https://img.shields.io/github/stars/elky84/ExcelToDotnet.svg?style=social&label=Stars)
![GitHub watchers](https://img.shields.io/github/watchers/elky84/ExcelToDotnet.svg?style=social&label=Watch)
![GitHub followers](https://img.shields.io/github/followers/elky84.svg?style=social&label=Follow)

![GitHub](https://img.shields.io/github/license/mashape/apistatus.svg)
![GitHub repo size in bytes](https://img.shields.io/github/repo-size/elky84/ExcelToDotnet.svg)
![GitHub code size in bytes](https://img.shields.io/github/languages/code-size/elky84/ExcelToDotnet.svg)

# ExcelToDotnet

## Nuget.org

<https://www.nuget.org/packages/ExcelToDotnet/>

## introduce

### English

Excel To Dotnet Compatible Data (Enum, Class, JSON)

It can be said to be a converter that can be used in Unity, C# applications, etc.

An Excel Sheet with a set rule is required.

In the case of Enum, only the Enum sheet must be registered. Otherwise, the sheet name becomes the class name.

In all cases, # is used as a comment (table, column, etc.).

In case of Enum, start :Begin and end point should be :End.

In the case of a table, the first row must be the column name, and the end point must be specified with :End.
The second row is the data type, and it is possible to link to the Id column of another table with $. 

### Korean

Unity, C# 애플리케이션 등에서 사용할 수 있는 변환기라고 할 수 있습니다.

규칙이 설정된 Excel 시트가 필요합니다.

Enum의 경우 Enum 시트만 등록해야 합니다. 그렇지 않으면 시트 이름이 클래스 이름이 됩니다.

모든 경우에 #은 주석(테이블, 열 등)으로 사용됩니다.

Enum의 경우 시작 :Begin, 끝점은 :End여야 합니다.

테이블의 경우 첫 번째 행은 열 이름이어야 하며 끝점은 :End로 지정해야 합니다.
두 번째 행은 데이터 타입으로 $로 다른 테이블의 Id 컬럼과 연결이 가능 합니다. 

## Sample Excel (xlsx)

<https://github.com/elky84/ExcelToDotnet/blob/main/ExcelCli/Character.xlsx>

## add package

`dotnet add package ExcelToDotnet`

## Implment CLI. (link ExcelToDotnet)

Release: <https://github.com/elky84/ExcelToDotnet/releases>

Reference : <https://github.com/elky84/ExcelToDotnet/blob/main/ExcelCli/Program.cs>, <https://github.com/elky84/ExcelToDotnet/blob/main/ExcelCli>

## install cli global tool

require dotnet 6 (LTS) or later (<https://dotnet.microsoft.com/en-us/download>)

`dotnet tool install -g ExcelCli`

## global tool usage

execute command name is `excel2dotnet`

### use single excel file (-f)
`excel2dotnet -f {fileName}`

### use target directory (-d)
`excel2dotnet -d {directory}`

### use enum generate mode (-e)
`excel2dotnet -d {directory} -e`

### use validation mode (-v)
`excel2dotnet d {directory} -v`

### use nullable mode (-l) => for .NET 6 or later
`excel2dotnet d {directory} -l`

## Execute CLI options (execute build file)

execute file name `excel2dotnet` instead of `ExcelCli`

### all options
- <https://github.com/elky84/ExcelToDotnet/blob/main/ExcelToDotnet/Options.cs>
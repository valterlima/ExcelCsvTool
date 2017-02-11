# ExcelCsvTool
ASP.NET MVC 5 App to convert Excel files to CSV

## Requirements
* [NuGet](https://www.nuget.org/)
* [.NET Framework v4.6.2](https://www.microsoft.com/en-us/download/details.aspx?id=53345)

## Dependencies
Download and install packages with NuGet: 

1. `nuget restore`

## How to build
To build you only need **MsBuild** from the .NET Framework

1. `MsBuild.exe .\CsvConverter.sln /t:Build /p:Configuration=Release /p:TargetFramework=v4.6.2`

## Run with IIS Express
1. `iisexpress.exe /path:"**your_path_to**\CsvConverter\CsvConverter"`

## Current validations
* File size must be < 1 MB. 
* File type must be Excel Office 2003- (.xls) or 2007+ (.xlsx).


# ExcelCsvTool
ASP.NET MVC 5 App to convert Excel files to CSV

## Requirements
* [NuGet](https://www.nuget.org/)
* [.NET Framework v4.6.2](https://www.microsoft.com/en-us/download/details.aspx?id=53345)

## Dependencies
Download and install packages with NuGet: 

<pre>nuget restore</pre>

## How to build
To build you only need **MsBuild** from the .NET Framework

<pre>MsBuild.exe .\CsvConverter.sln /t:Build /p:Configuration=Release /p:TargetFramework=v4.6.2</pre>

## Running

### Running with IIS Express
You need to have IIS Express installed. (It comes with Visual Studio)
<pre>iisexpress.exe /path:"<b>full_path</b>\CsvConverter" /clr:v4.0</pre>

Browse to http://localhost:8080 or something like that (depends on your config file)

### Deploy to azure
1. Fork this repository
2. Configure your Azure resource to use your forked repo

### Running with docker (Only Windows 10)

<pre>TODO</pre>

### Current validations
* File size must be < 1 MB. 
* File type must be Excel Office 2003- (.xls) or 2007+ (.xlsx).


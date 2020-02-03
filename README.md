# Author - Main project
https://github.com/marcinKotynia/ExcelToHtml

# ExcelToHtml.dll
Excel To HTML Library in .Net Standart 2.1

# List of Features (1.3)

- Convert Excel to HTML
	- Support for .xlsx format (Microsoft Office 2007+) 
	- Excel Properties: Border,border collor, Text-align, background-color, color, font-weight, font-size, width, white-space
	- Horizontal Merged Cells
	- Hidden Rows and columns
	- Comments
	- Injection safe
- Support for Functions  ( https://epplus.codeplex.com/wikipage?title=Supported%20Functions&referringTitle=Documentation )
- Calculation Engine
- Merge object, Json, REST API and excel template, convert to html

# Getting Started

## ExcelToHtml.dll, Nuget Package https://www.nuget.org/packages/ExcelToHtml

Basic Convert excel to HTML

```c#
FileInfo excelfile = new FileInfo(path);
var WorksheetHtml = new ExcelToHtml.ToHtml(excelfile);
string html = WorksheetHtml.GetHtml();
```

ExcelToHtml as calculation engine InputOutput

```c#
FileInfo newFile = new FileInfo(fullPath);
var WorksheetHtml =  new ExcelToHtml.ToHtml(ExcelFile);

//Optional Get Set Cells
Dictionary<string, string> InputOutput = new Dictionary<string, string>();
InputOutput.Add("A1", "Hello World");  			//set hello world
InputOutput.Add("A2", "=2+1");  			//set formula
InputOutput.Add("[[TemplateField]]", "HelloTemplate");  //FillTempalte Field
InputOutput.Add(".A2", null);  				//Output value form A2
var output = WorksheetHtml.GetSetCells(InputOutput);	//Output

string html = WorksheetHtml.Convert();
```

Merge REST API data and excel template, get html

```c#
FileInfo newFile = new FileInfo(fullPath);
var WorksheetHtml =  new ExcelToHtml.ToHtml(ExcelFile);
WorksheetHtml.DebugMode = true;
WorksheetHtml.DataFromUrl("http://nflarrest.com/api/v1/crime");
string html = WorksheetHtml.GetHtml();
```

Merge object and excel template, get html

```c#
FileInfo newFile = new FileInfo(fullPath);
var WorksheetHtml =  new ExcelToHtml.ToHtml(ExcelFile);
WorksheetHtml.DataFromObject(object); 
string html = WorksheetHtml.GetHtml();
```

Merge json and excel template, get html 

```c#
FileInfo newFile = new FileInfo(fullPath);
var WorksheetHtml =  new ExcelToHtml.ToHtml(ExcelFile);
WorksheetHtml.DataFromJson(string); 
string html = WorksheetHtml.GetHtml();
```

Merge json and excel template, get Excel

```c#
FileInfo newFile = new FileInfo(fullPath);
var WorksheetHtml =  new ExcelToHtml.ToHtml(ExcelFile);
WorksheetHtml.DataFromJson(string); 
Bytes[] html = WorksheetHtml.GetBytes();
```

# Technical Appendix

## List of Unsupported Features
- Vertical merged cells
- Charts 
- Images

## Colors and Themes
Getting color for a font, background is really challenging.
There are 3 different scenarios 

1. Themes (Supported only default theme)
2. System Colors with Index (supported)
3. RGB colors (supported)

This script will convert background color and font color to rgb colors if you use custom theme
and colours. To use 

1. open file in Excel 
2. Alt+F11 
3. Paste and Run code using F5

Result: Colors (background,borders,font) will be converted to RGB colors

```vb
Sub SheetBackgroundColorsToRgb()

Application.ScreenUpdating = False

    For Each Cell In ActiveSheet.UsedRange.Cells
    
		'Background
        Dim colorVal As Variant
        colorVal = Cell.Interior.Color
        Cell.Interior.Color = RGB((colorVal Mod 256), ((colorVal \ 256) Mod 256), (colorVal \ 65536))
        
        'Font color
        colorVal = Cell.Font.Color
        If (Not colorVal) Then
        Cell.Font.Color = RGB((colorVal Mod 256), ((colorVal \ 256) Mod 256), (colorVal \ 65536))
        End If
        
        'Borders     
        colorVal = Cell.Borders(xlEdgeBottom).Color
        If (Not colorVal) Then
        Cell.Borders(xlEdgeBottom).Color = RGB((colorVal Mod 256), ((colorVal \ 256) Mod 256), (colorVal \ 65536))
        End If
        
        colorVal = Cell.Borders(xlEdgeRight).Color
        If (Not colorVal) Then
        Cell.Borders(xlEdgeRight).Color = RGB((colorVal Mod 256), ((colorVal \ 256) Mod 256), (colorVal \ 65536))
        End If
        
        colorVal = Cell.Borders(xlEdgeTop).Color
        If (Not colorVal) Then
        Cell.Borders(xlEdgeTop).Color = RGB((colorVal Mod 256), ((colorVal \ 256) Mod 256), (colorVal \ 65536))
        End If
        
        colorVal = Cell.Borders(xlEdgeLeft).Color
        If (Not colorVal) Then
        Cell.Borders(xlEdgeLeft).Color = RGB((colorVal Mod 256), ((colorVal \ 256) Mod 256), (colorVal \ 65536))
        End If
        
    Next
    
Application.ScreenUpdating = True

End Sub
```

## Some technology remarks that could help you do even more :)
- xlsx file format is zip file with embeded xml files (https://en.wikipedia.org/wiki/Office_Open_XML )
- Libraries thet will help you
	- EPPlus http://epplus.codeplex.com/ 
	- Microsoft wrapper for handling openxml https://www.nuget.org/packages/DocumentFormat.OpenXml 
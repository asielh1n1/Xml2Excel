# Xml2Excel
Library to convert xml to excel using Net.Core

## How to use it?
### Install with NuGet
Package Manager
```sh
Install-Package Xml2Excel
```
.NET CLI
```sh
dotnet add package Xml2Excel --version 1.0.1
```
### Create an XML file with this structure
File: excel-test.xml
```sh
<workbook author="Asiel Hernandez Valdes" title="Excel Test">
	<worksheets>
		<worksheet name="tab1">
			<cells>
				<cell row="1" column="1" >Hello</cell>
				<cell row="1" column="2" >World</cell>
			</cells>
		</worksheet>
	</worksheets>
</workbook>
```
C# Code
```sh
Xml2ExcelCore xml2ExcelCore = new Xml2ExcelCore();
string xml = File.ReadAllText("excel-test.xml");
bool result = xml2ExcelCore.Generate(xml, "excel-file.xlsx");
```
Or
```sh
Xml2ExcelCore xml2ExcelCore = new Xml2ExcelCore();
string xml = File.ReadAllText("excel-test.xml");
using (MemoryStream memoryStream = xml2ExcelCore.Generate(xml))
{
	var buffer = memoryStream.ToArray();
	await File.WriteAllBytesAsync("excel-file.xlsx", buffer);
}
```
## XML tag properties
### workbook
Properties:
- title: 
- author 
- subject
- category
- keywords
- comments
- status
- company
- manager
- worksheets: Worksheet list

Example
```sh
<workbook 
   author="Asiel Hernandez Valdes" 
   title="Excel Test"
   subject="theSubject"
   category="theCategory"
   keywords="theKeywords"
   comments="theComments"
   status="theStatus"
   company="theCompany"
   manager="theManager"
>
	...
</workbook>
```

### Tag worksheet
Properties:
- name: Tab name
- tabColor: Allows to put colors in html format (#2196f3)
- rowHeight: Change the default row height for all new worksheets in this workbook
- password: Protected Password

Example
```sh
<workbook>
	<worksheets>
		<worksheet name="tab1" rowHeight="30">
			...
		</worksheet>
      <worksheet name="tab2" tabColor="#2196f3"  password="1234">
			...
		</worksheet>
	</worksheets>
</workbook>

```

### Tag range
Properties:
- cell1: Specify the start cell. It must be in the format "row, column"
- cell2: Specify the end cell. It must be in the format "row, column"
- merge: Merged cells. Possible values "true", "false"
- clear: Clear cells. Possible values "true", "false"

Exmaple
```sh
<workbook>
	<worksheets>
		<worksheet name="tab1">
			<ranges>
				<range cell1="4,1" cell2="4,4" merge="true" />
			</ranges>
         <cells>
				<cell row="4" column="1">Merge Cell</cell>
			</cells>
		</worksheet>
	</worksheets>
</workbook>
```

### Tag row
Properties:
- number: Specify number of row. Accepts only integer values
- height: Specify height of row. Accepts only integer values
- adjustToContents: Adjust Row Height to Contents. Possible values "true", "false"
- style: Html css style format(style="text-align:center;color:#ff0000")

Exmaple
```sh
<workbook>
	<worksheets>
		<worksheet name="tab1">
			<rows>
				<row number="7" height="30" style="text-align:center;background-color:#00ff00;color:#0000ff"/>
			</rows>
         ...
		</worksheet>
	</worksheets>
</workbook>
```

### Tag column
Properties:
- number: Specify number of row. Accepts only integer values
- width: Specify height of row. Accepts only integer values
- adjustToContents: Adjust Row Height to Contents. Possible values "true", "false"
- style: Html css style format(style="text-align:center;color:#ff0000")

Exmaple
```sh
<workbook>
	<worksheets>
		<worksheet name="tab1">
			<columns>
				<column number="7" width="50" style="text-align:center;background-color:#ff0000;color:#0000ff"></column>
			</columns>
         ...
		</worksheet>
	</worksheets>
</workbook>
```

### Tag cell
Properties:
- row: It must be an integer value starting with 1
- column: It must be an integer value starting with 1
- formula: Lets write formulas(Ex. A1+B1)
- link: Lets open link
- style: Html css style format(style="text-align:center;color:#ff0000")
- image: Allows you to incorporate an image. It is possible to use a url or a path
- imageScale: The "image" property is required, and allows the image to be scaled
- numberFormat: Using a custom number format. Example "$ #,##0.00"
- formatId:  Open XML offers predefined formats for dates and numbers. You can find them [here](https://github.com/closedxml/closedxml/wiki/NumberFormatId-Lookup-Table).
#### Supported styles:
- text-align:Possible values,"center","left","right"
- border-color: Allows to put colors in html format (#2196f3) and applies to all edges, top, right, bottom and left
- border-left-color: Allows to put colors in html format (#2196f3)
- border-right-color: Allows to put colors in html format (#2196f3)
- border-top-color: Allows to put colors in html format (#2196f3)
- border-bottom-color: Allows to put colors in html format (#2196f3)
- border-style:Possible values,"solid","dotted","dashed","none". Applies to all edges, top, right, bottom and left
- border-top-style: Possible values,"solid","dotted","dashed","none"
- border-right-style: Possible values,"solid","dotted","dashed","none"
- border-bottom-style: Possible values,"solid","dotted","dashed","none"
- border-left-style: Possible values,"solid","dotted","dashed","none"
- background-color: Allows to put colors in html format (#2196f3)
- background-color-pattern: Possible values,"solid","darktrellis","lighttrellis","none","darkhorizontal","lighthorizontal","darkvertical","lightvertical","darkdown","lightdown","darkup","lightup","lightgray","darkgray","darkgrid","lightgrid"
- font-style: Possible values "italic","shadow"
- font-weight: Possible values "bold","normal"
- text-decoration: Possible values "none", "underline","underline-double","strikethrough"
- font-size: Only accepts double values
- font-family: Example "Bahnschrift"

Example:
```sh
<workbook>
	<worksheets>
		<worksheet name="Tab test" tabColor="#2196f3">
			<cells>
				<cell row="1" column="1" >2</cell>
				<cell row="1" column="2" >3</cell>
				<cell row="1" column="3" formula="A1+B1"></cell>
            <cell row="3" column="2" style="text-align:center;background-color:#ff0000;color:#0000ff">Cell Test</cell>
			</cells>
		</worksheet>
	</worksheets>
</workbook>
```


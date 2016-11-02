# VBA Project: **clsmPivotTables**
## VBA Module: **[UDFs](/libraries/UDFs.vba "source is here")**
### Type: StdModule  

This procedure list for repo (clsmPivotTables) was automatically created on 02.11.2016 13:36:14 by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in UDFs

---
VBA Procedure: **LogikVBA**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function LogikVBA(v) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
v|Variant|False||


---
VBA Procedure: **VLOOKUP3**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function VLOOKUP3(Table As Range, SearchColumnNum As Integer, SearchValue As Variant, ResultColumnNum As Integer)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
Table|Range|False||
SearchColumnNum|Integer|False||
SearchValue|Variant|False||
ResultColumnNum|Integer|False||


---
VBA Procedure: **Translit**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function Translit(txt As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
txt|String|False||


---
VBA Procedure: **MultiCat**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function MultiCat(ByRef rng As Excel.Range, Optional ByVal DELIM As String = "") As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByRef|Excel|False||
ByVal|Variant|True||


---
VBA Procedure: **Lotto**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function Lotto(Bottom As Integer, Top As Integer, Amount As Integer)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
Bottom|Integer|False||
Top|Integer|False||
Amount|Integer|False||


---
VBA Procedure: **RandomSelect**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function RandomSelect(TargetCells)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
TargetCells|Variant|False||


---
VBA Procedure: **WeekdayWord**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function WeekdayWord(MyDate As Date) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
MyDate|Date|False||


---
VBA Procedure: **Class**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function Class(m, i)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
m|Variant|False||
i|Variant|False||


---
VBA Procedure: **PropisRus**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function PropisRus(n As Double, rub As Boolean) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
n|Double|False||
rub|Boolean|False||


---
VBA Procedure: **PropisEng**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function PropisEng(ByVal strAmount As String, strCur As String, strDec As String, iPrec As Integer)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|String|False||
strCur|String|False||
strDec|String|False||
iPrec|Integer|False||


---
VBA Procedure: **GetHundreds**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function GetHundreds(ByVal MyNumber)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|Variant|False||


---
VBA Procedure: **GetTens**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function GetTens(TensText)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
TensText|Variant|False||


---
VBA Procedure: **GetDigit**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function GetDigit(Digit)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
Digit|Variant|False||


---
VBA Procedure: **SumBetween**  
Type: **Function**  
Returns: **Long**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function SumBetween(TargetCells As Range, min As Long, max As Long, IncludeMin As Boolean, IncludeMax As Boolean) As Long*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
TargetCells|Range|False||
min|Long|False||
max|Long|False||
IncludeMin|Boolean|False||
IncludeMax|Boolean|False||


---
VBA Procedure: **NeedDate**  
Type: **Function**  
Returns: **Date**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function NeedDate(n As Integer, w As Integer, m As Integer, Y As Integer) As Date*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
n|Integer|False||
w|Integer|False||
m|Integer|False||
Y|Integer|False||


---
VBA Procedure: **GetNumbers**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function GetNumbers(TargetCell As Range) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
TargetCell|Range|False||


---
VBA Procedure: **GetText**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function GetText(TargetCell As Range) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
TargetCell|Range|False||


---
VBA Procedure: **FirstInRow**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function FirstInRow(myRow As Range)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
myRow|Range|False||


---
VBA Procedure: **FirstInColumn**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function FirstInColumn(myColumn As Range)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
myColumn|Range|False||


---
VBA Procedure: **LastInRow**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function LastInRow(myRow As Range)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
myRow|Range|False||


---
VBA Procedure: **LastInColumn**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function LastInColumn(myColumn As Range)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
myColumn|Range|False||


---
VBA Procedure: **SheetName1**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function SheetName1() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **WorkbookName**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function WorkbookName() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **FullFileName**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function FullFileName() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **UserName**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function UserName() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **CellColor**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function CellColor(cell As Range)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
cell|Range|False||


---
VBA Procedure: **CellFontColor**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function CellFontColor(cell As Range)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
cell|Range|False||


---
VBA Procedure: **AutoFilter_Criteria**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function AutoFilter_Criteria(Header As Range) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
Header|Range|False||


---
VBA Procedure: **Substring**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function Substring(txt, Delimiter, n) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
txt|Variant|False||
Delimiter|Variant|False||
n|Variant|False||


---
VBA Procedure: **GrupString**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function GrupString() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **VLOOKUP2**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function VLOOKUP2(Table As Range, SearchColumnNum As Integer, SearchValue As Variant, n As Integer, ResultColumnNum As Integer)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
Table|Range|False||
SearchColumnNum|Integer|False||
SearchValue|Variant|False||
n|Integer|False||
ResultColumnNum|Integer|False||


---
VBA Procedure: **MaskCompare**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function MaskCompare(txt As String, Mask As String, CaseSensitive As Boolean)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
txt|String|False||
Mask|String|False||
CaseSensitive|Boolean|False||


---
VBA Procedure: **CountByMask**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function CountByMask(rng As Range, Mask As String, CaseSensitive As Boolean)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
rng|Range|False||
Mask|String|False||
CaseSensitive|Boolean|False||


---
VBA Procedure: **IsLatin**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function IsLatin(txt As String)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
txt|String|False||


---
VBA Procedure: **SumByCellColor**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function SumByCellColor(SearchRange As Range, TargetCell As Range)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
SearchRange|Range|False||
TargetCell|Range|False||


---
VBA Procedure: **SumByFontColor**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function SumByFontColor(SearchRange As Range, TargetCell As Range)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
SearchRange|Range|False||
TargetCell|Range|False||


---
VBA Procedure: **MicroCharts**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function MicroCharts(rng As Range)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
rng|Range|False||


---
VBA Procedure: **LastRow**  
Type: **Function**  
Returns: **Long**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function LastRow(SheetName As String) As Long*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
SheetName|String|False||


---
VBA Procedure: **LastColumn**  
Type: **Function**  
Returns: **Range**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function LastColumn(SheetName As String, r As Long) As Range*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
SheetName|String|False||
r|Long|False||


---
VBA Procedure: **SheetExist**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function SheetExist(SheetName As String) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
SheetName|String|False||


---
VBA Procedure: **SheetExistBook**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function SheetExistBook(wb As Workbook, SheetName As String) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
wb|Workbook|False||
SheetName|String|False||


---
VBA Procedure: **SheetExistBookCreate**  
Type: **Function**  
Returns: **Worksheet**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function SheetExistBookCreate(wb As Workbook, SheetName, cl As Boolean) As Worksheet*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
wb|Workbook|False||
SheetName|Variant|False||
cl|Boolean|False||


---
VBA Procedure: **InversiaValue**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function InversiaValue(v As Range)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
v|Range|False||


---
VBA Procedure: **Delimeter_Count**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function Delimeter_Count(r As Range, Delimeter As String)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
r|Range|False||
Delimeter|String|False||


---
VBA Procedure: **SelectFiles**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function SelectFiles(MultiSelect As Boolean, fname As String, f As String)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
MultiSelect|Boolean|False||
fname|String|False||
f|String|False||


---
VBA Procedure: **ИмяЛистаВАпостров**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function ИмяЛистаВАпостров(Имя As String)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
As|Variant|False||


---
VBA Procedure: **ЗаголовокСтолбца**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function ЗаголовокСтолбца(wb As Workbook, Sh As String, NameZag As String) As Range*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
wb|Workbook|False||
Sh|String|False||
NameZag|String|False||


---
VBA Procedure: **ЗаголовокСтолбцаСоздатьИлиВернуть**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function ЗаголовокСтолбцаСоздатьИлиВернуть(wb As Workbook, sSH_Name As String, sNameZag As String)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
wb|Workbook|False||
sSH_Name|String|False||
sNameZag|String|False||


---
VBA Procedure: **grt**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub grt()*  

**no arguments required for this procedure**


---
VBA Procedure: **CommentExist**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function CommentExist(r As Range) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
r|Range|False||


---
VBA Procedure: **CommentTEXT**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function CommentTEXT(r As Range)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
r|Range|False||


---
VBA Procedure: **WorkbookExist**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function WorkbookExist(Path, WorkbookName As String) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
Path|Variant|False||
WorkbookName|String|False||


---
VBA Procedure: **IsBookOpen**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function IsBookOpen(wbFullName As String) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
wbFullName|String|False||


---
VBA Procedure: **ОТСТУП**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function ОТСТУП(ячейка)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
|Variant|False||

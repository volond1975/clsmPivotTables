# VBA Project: **clsmPivotTables**
## VBA Module: **[clsmPivotTables](/scripts/clsmPivotTables.cls "source is here")**
### Type: ClassModule  

This procedure list for repo (clsmPivotTables) was automatically created on 02.11.2016 13:36:14 by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in clsmPivotTables

---
VBA Procedure: **sbInitError**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Sub sbInitError()*  

**no arguments required for this procedure**


---
VBA Procedure: **NewEnum**  
Type: **Get**  
Returns: **IUnknown**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get NewEnum() As IUnknown*  

**no arguments required for this procedure**


---
VBA Procedure: **Initialize**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function Initialize(WbWithTables As Excel.Workbook)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
WbWithTables|Excel|False||


---
VBA Procedure: **Refresh**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Sub Refresh()*  

**no arguments required for this procedure**


---
VBA Procedure: **Item**  
Type: **Get**  
Returns: **ListObject**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get Item(Index As Variant) As Excel.ListObject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
Index|Variant|False||


---
VBA Procedure: **Item_PivotCache**  
Type: **Get**  
Returns: **PivotCache**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get Item_PivotCache(Index As Variant) As Excel.PivotCache*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
Index|Variant|False||


---
VBA Procedure: **Count**  
Type: **Get**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get Count()*  

**no arguments required for this procedure**


---
VBA Procedure: **Count_PivotCaches**  
Type: **Get**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get Count_PivotCaches()*  

**no arguments required for this procedure**


---
VBA Procedure: **Exists**  
Type: **Get**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Property Get Exists(Index As Variant) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
Index|Variant|False||


---
VBA Procedure: **SheetExists**  
Type: **Get**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Property Get SheetExists(Index As Variant) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
Index|Variant|False||


---
VBA Procedure: **Items**  
Type: **Get**  
Returns: **Collection**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get Items() As Collection*  

**no arguments required for this procedure**


---
VBA Procedure: **Items_PivotCaches**  
Type: **Get**  
Returns: **Collection**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get Items_PivotCaches() As Collection*  

**no arguments required for this procedure**


---
VBA Procedure: **Workbook**  
Type: **Let**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Property Let Workbook(ByVal sFullNameBook As String)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|String|False||


---
VBA Procedure: **Workbook**  
Type: **Get**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Property Get Workbook() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **Worksheet**  
Type: **Let**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Property Let Worksheet(ByVal sSheetName As String)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|String|False||


---
VBA Procedure: **Worksheet**  
Type: **Get**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Property Get Worksheet() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **name**  
Type: **Let**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Property Let name(ByVal sNameListObject As String)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|String|False||


---
VBA Procedure: **name**  
Type: **Get**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Property Get name() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **NameStyle**  
Type: **Let**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Property Let NameStyle(ByVal sNameStyleTable As String)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|String|False||


---
VBA Procedure: **NameStyle**  
Type: **Get**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Property Get NameStyle() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **RangeStr**  
Type: **Let**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Property Let RangeStr(ByVal sRange As String)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|String|False||


---
VBA Procedure: **RangeStr**  
Type: **Get**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Property Get RangeStr() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **RangeRng**  
Type: **Let**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Property Let RangeRng(ByVal rngRange As Range)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|Range|False||


---
VBA Procedure: **RangeRng**  
Type: **Get**  
Returns: **Range**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Property Get RangeRng() As Range*  

**no arguments required for this procedure**


---
VBA Procedure: **fnGetNormaliseAddr**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Friend Function fnGetNormaliseAddr(ByVal sRng As String, Optional bMulti As Boolean = False) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|String|False||
bMulti|Boolean|True| False|


---
VBA Procedure: **fnAreaType**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function fnAreaType(RangeArea As Range) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
RangeArea|Range|False||


---
VBA Procedure: **fnCheckShablon**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Friend Function fnCheckShablon(ByVal sRng As String) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|String|False||


---
VBA Procedure: **fnGetFileName**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Friend Function fnGetFileName(ByVal strFullNameFile As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|String|False||


---
VBA Procedure: **fnGetFileExt**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Friend Function fnGetFileExt(ByVal strNameFile As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|String|False||


---
VBA Procedure: **fnGetFilePath**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Friend Function fnGetFilePath(ByVal strFullNameFile As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|String|False||


---
VBA Procedure: **fnGetPathLevel**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Friend Function fnGetPathLevel(ByVal strFullPath As String, Optional ByVal iLvl As Integer = 9999) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|String|False||
ByVal|Variant|True||


---
VBA Procedure: **sbSetError**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Sub sbSetError(Optional ByVal lParamErr As Long = 0, Optional ByVal sErrText As String = "", Optional ByVal sReplaceParam As String = "")*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|Variant|True||
ByVal|Variant|True||
ByVal|Variant|True||


---
VBA Procedure: **GetError**  
Type: **Get**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Property Get GetError()*  

**no arguments required for this procedure**


---
VBA Procedure: **IsError**  
Type: **Get**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Property Get IsError() As Boolean*  

**no arguments required for this procedure**


---
VBA Procedure: **Header**  
Type: **Let**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Property Let Header(ByVal bSetHeader As Boolean)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|Boolean|False||


---
VBA Procedure: **Header**  
Type: **Get**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Property Get Header() As Boolean*  

**no arguments required for this procedure**


---
VBA Procedure: **sbCreateDefValue**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub sbCreateDefValue()*  

**no arguments required for this procedure**


---
VBA Procedure: **TABLEDATARANGE**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function TABLEDATARANGE(ByVal loTest As ListObject) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|ListObject|False||


---
VBA Procedure: **TABLEHEADERRANGE**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function TABLEHEADERRANGE(ByVal loTest As ListObject) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|ListObject|False||


---
VBA Procedure: **TABLETOTALSRANGE**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function TABLETOTALSRANGE(ByVal loTest As ListObject) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|ListObject|False||


---
VBA Procedure: **TABLEHEADERVISIBLE**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function TABLEHEADERVISIBLE(ByVal TableCell As Range) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|Range|False||


---
VBA Procedure: **TABLETOTALSROWVISIBLE**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function TABLETOTALSROWVISIBLE(ByVal TableCell As Range) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|Range|False||


---
VBA Procedure: **add_PivotCache**  
Type: **Get**  
Returns: **PivotCache**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Property Get add_PivotCache(Optional ByVal SourceType = 1, Optional SourceData) As Excel.PivotCache*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|Variant|True||
SourceData|Variant|True||


---
VBA Procedure: **CreatePivotCache**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function CreatePivotCache(Optional SourceType = 1, Optional SourceData)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
SourceType|Variant|True||
SourceData|Variant|True||


---
VBA Procedure: **CreatePivotTable2**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub CreatePivotTable2(PTcache As PivotCache)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
PTcache|PivotCache|False||


---
VBA Procedure: **Add_PivotTable**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function Add_PivotTable(pvtCache As PivotCache, TableName, Optional TableDestination)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
pvtCache|PivotCache|False||
TableName|Variant|False||
TableDestination|Variant|True||


---
VBA Procedure: **CreatePivotsb**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub CreatePivotsb()*  

**no arguments required for this procedure**


---
VBA Procedure: **ShowCacheIndex**  
Type: **Function**  
Returns: **Long**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function ShowCacheIndex(rngPT As Range) As Long*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
rngPT|Range|False||


---
VBA Procedure: **GetMemory**  
Type: **Function**  
Returns: **Long**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function GetMemory(rngPT As Range) As Long*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
rngPT|Range|False||


---
VBA Procedure: **GetRecords**  
Type: **Function**  
Returns: **Long**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function GetRecords(rngPT As Range) As Long*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
rngPT|Range|False||


---
VBA Procedure: **ChangePivotCache**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub ChangePivotCache()*  

**no arguments required for this procedure**


---
VBA Procedure: **SelPTNewCache**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub SelPTNewCache()*  

**no arguments required for this procedure**


---
VBA Procedure: **CheckCaches**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub CheckCaches()*  

**no arguments required for this procedure**


---
VBA Procedure: **DeleteOldItemsWB**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub DeleteOldItemsWB()*  

**no arguments required for this procedure**


---
VBA Procedure: **DeleteAllPivotTables**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub DeleteAllPivotTables()*  

**no arguments required for this procedure**


---
VBA Procedure: **DeletePivotTable**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub DeletePivotTable()*  

**no arguments required for this procedure**


---
VBA Procedure: **Adding_PivotFields**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub Adding_PivotFields()*  

**no arguments required for this procedure**


---
VBA Procedure: **AddCalculatedField**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub AddCalculatedField()*  

**no arguments required for this procedure**


---
VBA Procedure: **AddValuesField**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub AddValuesField()*  

**no arguments required for this procedure**


---
VBA Procedure: **RemovePivotField**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub RemovePivotField()*  

**no arguments required for this procedure**


---
VBA Procedure: **RemoveCalculatedField**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub RemoveCalculatedField()*  

**no arguments required for this procedure**


---
VBA Procedure: **ReportFiltering_Single**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub ReportFiltering_Single()*  

**no arguments required for this procedure**


---
VBA Procedure: **ReportFiltering_Multiple**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub ReportFiltering_Multiple()*  

**no arguments required for this procedure**


---
VBA Procedure: **ClearReportFiltering**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub ClearReportFiltering()*  

**no arguments required for this procedure**


---
VBA Procedure: **RefreshingPivotTables**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub RefreshingPivotTables()*  

**no arguments required for this procedure**


---
VBA Procedure: **ChangePivotDataSourceRange**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub ChangePivotDataSourceRange()*  

**no arguments required for this procedure**


---
VBA Procedure: **PivotGrandTotals**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub PivotGrandTotals(Index As Variant)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
Index|Variant|False||


---
VBA Procedure: **PivotReportLayout**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub PivotReportLayout(Index As Variant)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
Index|Variant|False||


---
VBA Procedure: **PivotTable_DataFormatting**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub PivotTable_DataFormatting(Index As Variant)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
Index|Variant|False||


---
VBA Procedure: **PivotField_DataFormatting**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub PivotField_DataFormatting(Index As Variant, PivotFieldName)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
Index|Variant|False||
PivotFieldName|Variant|False||


---
VBA Procedure: **PivotField_ExpandCollapse**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub PivotField_ExpandCollapse(Index As Variant, PivotFieldName)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
Index|Variant|False||
PivotFieldName|Variant|False||


---
VBA Procedure: **RepeatLabels**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub RepeatLabels(Index As Variant)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
Index|Variant|False||


---
VBA Procedure: **CurrentPageSelect**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub CurrentPageSelect(Index As Variant, CurrentPageName, Optional CurrentPageValue = "(All)", Optional bEnableMultiplePageItems = False, Optional bRepeatLabels = True)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
Index|Variant|False||
CurrentPageName|Variant|False||
CurrentPageValue|Variant|True||
bEnableMultiplePageItems|Variant|True||
bRepeatLabels|Variant|True||

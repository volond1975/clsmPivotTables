Dim m_wb       As Excel.Workbook
Dim m_Tables   As Collection
Dim m_Tables_PivotCaches   As Collection
Dim m_Tables_ListColumns As Collection
Dim m_Tables_ListRows As Collection
Dim m_ValueListObject As Range
Dim m_pc As Excel.PivotCache
'===============================================
'ВЕРСИЯ КЛАССА
Private Const clsmListObjser = "0.0 (18.06.2015)"

'тип для перечня ошибок
Private Type tError
    NoneErr    As String
    CreateADOErr As String
    NoData     As String
    DataNotFound As String
    QueryError As String
    ConnectError As String
    DBConnectError As String
    WorkBookErr As String
    MultyRange As String
End Type


'внутренние объекты класса
Private sDataSource As String   'имя книги
Private sSheetSource As String  'имя листа
Private sNameSource As String  'имя таблицы
Private sNameStyle As String  'имя стиля таблицы

Private sActiveListObjectName As String
'рабочие пременные
Private iExcelVer As Integer    'версия EXCEL

'переключатели
Private bHeader As Boolean      'требует переподключения
Private bSubTotal As Boolean      'требует переподключения
'процедура, доступна только в этом классе
'формируем описание возможных ошибок
' %1 - меняется на текст подстановки
Private Sub sbInitError()
    eError.NoneErr = ""
    eError.CreateADOErr = "КРИТИЧЕСКАЯ ОШИБКА! НЕВОЗМОЖНО СОЗДАТЬ ОБЪЕКТЫ ADO"
    eError.NoData = "Ошибка. Запрос не создан"
    eError.DataNotFound = "Нет записей удоволетворяющих запросу '%1'"
    eError.QueryError = "Ошибка запроса: '%1'"
    eError.ConnectError = "Нет подключения к DB (сперва надо подключиться к '%1' используя метод BookConnect)"
    eError.DBConnectError = "Ошибка подключения к DB (проверьте правильность пути и наличия файла '%1')"
    eError.WorkBookErr = "Файл данных '%1' отсутствует"
    eError.MultyRange = "Задан множественный диапазон '%1' запрос не возможен"
End Sub

Public Property Get NewEnum() As IUnknown
'the following line, added in a text editor,
'creates the ability to cycle through the items with For Each
    Set NewEnum = m_Tables.[_NewEnum]
End Property

Public Function Initialize(WbWithTables As Excel.Workbook)
'Call sbInitError
    Call sbCreateDefValue

    Set m_wb = WbWithTables
    Call Refresh

End Function

Public Sub Refresh()
    Dim ws     As Excel.Worksheet
    Dim pt     As Excel.PivotTable
Dim pc     As Excel.PivotCache
    Set m_Tables = New Collection
    For Each ws In m_wb.Worksheets
        For Each pt In ws.PivotTables
            m_Tables.Add pt, pt.name
        Next pt
    Next ws
Set m_Tables_PivotCaches = New Collection
For Each pc In m_wb.PivotCaches

m_Tables_PivotCaches.Add pc, "pc_" & pc.Index
Next
End Sub

Public Property Get Item(Index As Variant) As Excel.ListObject
'the following line, added in a text editor,
'sets Item as the default property of the class
    Set Item = m_Tables(Index)
End Property

Public Property Get Item_PivotCache(Index As Variant) As Excel.PivotCache
'Index-number Index PivotCache
'the following line, added in a text editor,
'sets Item as the default property of the class
    Set Item_PivotCache = m_Tables_PivotCaches("pc_" & Index)
End Property

Public Property Get Count()
    Count = m_Tables.Count
End Property
Public Property Get Count_PivotCaches()
    Count_PivotCaches = m_Tables_PivotCaches.Count
End Property


Property Get Exists(Index As Variant) As Boolean
    Dim Test   As Variant
    On Error Resume Next
    Set Test = m_Tables(Index)
    Exists = Err.Number = 0
End Property
Property Get SheetExists(Index As Variant) As Boolean
    Dim Test   As Variant
    On Error Resume Next
    Set Test = m_wb.Worksheets(Index)
    SheetExists = Err.Number = 0
End Property



Public Property Get Items() As Collection

    Set Items = m_Tables
End Property

Public Property Get Items_PivotCaches() As Collection

    Set Items_PivotCaches = m_Tables_PivotCaches
End Property
'---------------------------------------------------------
'установить для класса рабочую книгу
'возможно автоматическое переподключение
Property Let Workbook(ByVal sFullNameBook As String)
    Dim bOk    As Boolean
    If sDataSource = sFullNameBook Then Exit Property
    If Dir(sFullNameBook) <> "" Then
        sDataSource = sFullNameBook
        '        bOk = fnReconnect()
    Else
        sbSetError sErrText:=eError.WorkBookErr, sReplaceParam:=sFullNameBook
    End If
End Property

'показать установленную для класса рабочую книгу
Property Get Workbook() As String
    Workbook = sDataSource
End Property

'установить для класса рабочий лист
Property Let Worksheet(ByVal sSheetName As String)
    If sSheetSource = sSheetName Then Exit Property
    sSheetSource = sSheetName
End Property

'показать текущий для класса рабочий лист
Property Get Worksheet() As String
    Worksheet = sSheetSource
End Property
'NameListObject

'установить для класса рабочий лист
Property Let name(ByVal sNameListObject As String)
    If sNameSource = sNameListObject Then Exit Property
    sNameSource = sNameListObject
End Property

'показать текущий для класса рабочий лист
Property Get name() As String
    name = sNameSource
End Property







'Property Get ActiveListObjectName() As String
'    ActiveListObjectName = sActiveListObjectName
'End Property




'установить для класса стиль таблицы
Property Let NameStyle(ByVal sNameStyleTable As String)
    If sNameStyle = sNameStyleTable Then Exit Property
    sNameStyle = sNameListObjectTable
End Property

'показать текущий для класса стиль таблицы
Property Get NameStyle() As String
    NameStyle = sNameStyle
End Property














'установить для класса нужную область на листе
'sRange - строка содержащая адрес
'например   "A" или "DF57" или "G3:EJ500" ...
'всегда нормализуется к диапазону
Property Let RangeStr(ByVal sRange As String)

    If sRange <> "" Then
        sRange = fnGetNormaliseAddr(sRange, True)
        eLastChange = eLastChange Or eRange    'пользователь меняет диапазон
    Else
        eLastChange = eLastChange Xor eRange    'пользователь отменил диапазон
    End If
    sUSRRange = sRange
End Property

'посмотреть (вернуть) заданную для класса область на листе
Property Get RangeStr() As String
    RangeStr = sUSRRange
End Property

'установить для класса нужную область на листе
'параметр - область
'напрмер Sheets("Info").Range("A23:D784") или Range("В3:B8") ...
Property Let RangeRng(ByVal rngRange As Range)
    Dim sRange As String, sBook As String, sBookPath As String
    'если заданы не связанные области - выходим
    If rngRange.Areas.Count > 1 Then
        sbSetError 0, eError.MultyRange, rngRange.Address
        Exit Property
    End If
    sRange = rngRange.Address

    'ПРОВЕРЯЕМ НА СООТВЕТСТВИЕ ЛИСТА И КНИГИ
    sSheetSource = rngRange.Parent.name
    sBook = rngRange.Parent.Parent.name
    sBookPath = rngRange.Parent.Parent.Path

    sRange = Replace(sRange, "$", "")
    eLastChange = eLastChange Or eRange
    sUSRRange = sRange

    If fnGetFileName(sDataSource) <> sBook Then
    End If

End Property

'посмотреть (вернуть) заданную для класса область на листе
Property Get RangeRng() As Range
    Dim sBookName As String
    sBookName = fnGetFileName(sDataSource)
    Set RangeRng = Workbooks(sBookName).Sheets(sSheetSource).Range(sUSRRange)
End Property
'ФУНКЦИЯ!!!! - доступна тут и у наследников
'возвращает "правильно" сформированный адрес в формате А1
'если установлено bMulti - возвращает в виде диапазона
' fnGetNormaliseAddr("A") вернет "A1"
' fnGetNormaliseAddr("A", True) вернет "A:A"
' fnGetNormaliseAddr("A1", True) вернет "A1:A1"
' fnGetNormaliseAddr("Af233:bb13", False) вернет "AF13"
' fnGetNormaliseAddr("Af233:bb13", True) вернет "Af233:bb13"
' fnGetNormaliseAddr("Af233:6bb13", False) вернет "AF233"
' fnGetNormaliseAddr("Af233:6bb13", True) вернет "Af233:Af233"   <<<<< 6bb13 - не верно заданный адрес
Friend Function fnGetNormaliseAddr(ByVal sRng As String, Optional bMulti As Boolean = False) As String
    Dim iPos As Integer, lCnt As Long, tmpRng As Range
    fnGetNormaliseAddr = "A1"

    'убираем знаки доллара и пробелы
    sRng = Replace(sRng, "$", ""): sRng = Trim(sRng)

    'если в диапазоне что-то указано
    If sRng <> "" Then
        iPos = InStr(sRng, ":")
        ' ?????
        If iPos = 0 Then
            If fnCheckShablon(sRng) Then
                sRng = sRng & ":" & sRng
            Else
                sRng = "A1:A1"
            End If
            ' *:*
        Else
            ' ?????:
            If iPos = Len(sRng) Then
                'проверяем на соответствие шаблону 1 часть адреса
                If fnCheckShablon(Mid(sRng, 1, iPos - 1)) Then
                    sRng = sRng & Mid(sRng, 1, iPos - 1)
                Else
                    sRng = "A1:A1"
                End If
                ' ????:????
            Else
                'проверяем на соответствие шаблону 1 часть адреса
                If Not fnCheckShablon(Mid(sRng, 1, iPos - 1)) Then
                    sRng = "A1" & Mid(sRng, iPos)
                    iPos = 3
                End If
                'проверяем на соответствие шаблону 2 часть адреса
                If Not fnCheckShablon(Mid(sRng, iPos + 1)) Then
                    sRng = Mid(sRng, 1, iPos) & Mid(sRng, 1, iPos - 1)
                End If
            End If
        End If
        'sRng - в нем ни чего нет - пусто
    Else
        sRng = "A1:A1"
    End If
    'проверим на правильность формирования диапазона
    'путем попытки подсчета в нем кол-ва ячеек
    On Error Resume Next
    lCnt = Range(sRng).Count
    'ошибка! что-то не так :( - сформируем из 1 части
    If Err <> 0 Then
        sRng = Mid(sRng, 1, iPos) & Mid(sRng, 1, iPos - 1)
    End If
    On Error GoTo 0
    Err.Clear
    'если не задан режим области ("мульти"), сформируем
    ' из области sRng - адрес левой верхней ячейки
    If Not bMulti Then
        sRng = Range(sRng).Cells(1, 1).Address
    End If
    'вернем нормализованный адрес
    fnGetNormaliseAddr = Replace(sRng, "$", "")
End Function


'определить тип области
Private Function fnAreaType(RangeArea As Range) As String
'   Возврат типа диапазона в области
    Select Case True
    Case RangeArea.Cells.CountLarge = 1
        AreaType = "Ячейка"
    Case RangeArea.CountLarge = Cells.CountLarge
        AreaType = "Рабочий лист"
    Case RangeArea.Rows.Count = Cells.Rows.Count
        AreaType = "Столбец"
    Case RangeArea.Columns.Count = Cells.Columns.Count
        AreaType = "Строка"
    Case Else
        AreaType = "Блок"
    End Select
End Function
'---------------------------------------------------------
'---------------------------------------------------------
'ФУНКЦИЯ!!!! - доступна тут и у наследников
'проверить строчку на соответствие шаблону адреса стиля А1
'сначала буквы (латынь) максимум 3 штуки, затем цифры максимкм 7 цифр
'если соответствует шаблону - возвращает True
Friend Function fnCheckShablon(ByVal sRng As String) As Boolean
    fnCheckShablon = False
    'удаляем знаки доллара и пробелы
    sRng = Replace(sRng, "$", ""): sRng = Trim(sRng)

    'A - ZZZ (указаны только буквы)
    If sRng Like "[A-Za-z]" Or _
       sRng Like "[A-Za-z][A-Za-z]" Or _
       sRng Like "[A-Za-z][A-Za-z][A-Za-z]" Then
        fnCheckShablon = True
        Exit Function
    End If
    'A1 - Z12345678  (указана буква и цифра/ы)
    If sRng Like "[A-Za-z][0-9]" Or _
       sRng Like "[A-Za-z][0-9][0-9]" Or _
       sRng Like "[A-Za-z][0-9][0-9][0-9]" Or _
       sRng Like "[A-Za-z][0-9][0-9][0-9][0-9]" Or _
       sRng Like "[A-Za-z][0-9][0-9][0-9][0-9][0-9]" Or _
       sRng Like "[A-Za-z][0-9][0-9][0-9][0-9][0-9][0-9]" Or _
       sRng Like "[A-Za-z][0-9][0-9][0-9][0-9][0-9][0-9][0-9]" Then
        fnCheckShablon = True
        Exit Function
    End If
    'AZ1 - AZ12345678 (2 буквы и цифры)
    If sRng Like "[A-Za-z][A-Za-z][0-9]" Or _
       sRng Like "[A-Za-z][A-Za-z][0-9][0-9]" Or _
       sRng Like "[A-Za-z][A-Za-z][0-9][0-9][0-9]" Or _
       sRng Like "[A-Za-z][A-Za-z][0-9][0-9][0-9][0-9]" Or _
       sRng Like "[A-Za-z][A-Za-z][0-9][0-9][0-9][0-9][0-9]" Or _
       sRng Like "[A-Za-z][A-Za-z][0-9][0-9][0-9][0-9][0-9][0-9]" Or _
       sRng Like "[A-Za-z][A-Za-z][0-9][0-9][0-9][0-9][0-9][0-9][0-9]" Then
        fnCheckShablon = True
        Exit Function
    End If
    'ABZ1 - ZZZ12345678 (3 буквы и цифры)
    If sRng Like "[A-Za-z][A-Za-z][A-Za-z][0-9]" Or _
       sRng Like "[A-Za-z][A-Za-z][A-Za-z][0-9][0-9]" Or _
       sRng Like "[A-Za-z][A-Za-z][A-Za-z][0-9][0-9][0-9]" Or _
       sRng Like "[A-Za-z][A-Za-z][A-Za-z][0-9][0-9][0-9][0-9]" Or _
       sRng Like "[A-Za-z][A-Za-z][A-Za-z][0-9][0-9][0-9][0-9][0-9]" Or _
       sRng Like "[A-Za-z][A-Za-z][A-Za-z][0-9][0-9][0-9][0-9][0-9][0-9]" Or _
       sRng Like "[A-Za-z][A-Za-z][A-Za-z][0-9][0-9][0-9][0-9][0-9][0-9][0-9]" Then
        fnCheckShablon = True
    End If
End Function

'---------------------------------------------------------
'---------------------------------------------------------
'ФУНКЦИЯ!!!! - доступна тут и у наследников
'fnGetFileName - вернуть имя файла вместе с расширением или ""
'pFulNameFile - имя файла (может содержать полный путь) (по значению)
Friend Function fnGetFileName(ByVal strFullNameFile As String) As String
    Dim iDirPos As Integer: iDirPos = 0
    Dim strPS  As String: strPS = Application.PathSeparator

    fnGetFileName = Mid(strFullNameFile, InStrRev(strFullNameFile, strPS) + 1)

End Function



'ФУНКЦИЯ!!!! - доступна тут и у наследников
'fnGetFileExt - вернуть расширение файла вместе с точкой или ""
'pNameFile - имя файла (может содержать несколько точек, обрабатывается последняя) (по значению)
Friend Function fnGetFileExt(ByVal strNameFile As String) As String
    Dim iDotPos As Integer: iDotPos = 0
    Dim strRetval As String: strRetval = ""

    iDotPos = InStrRev(strNameFile, ".")
    If iDotPos Then strRetval = LCase(Mid(strNameFile, iDotPos))

    fnGetFileExt = strRetval
End Function



'ФУНКЦИЯ!!!! - доступна тут и у наследников
'fnGetFilePath - вернуть путь к файлу
'pFulNameFile - имя файла включая путь (по значению)
Friend Function fnGetFilePath(ByVal strFullNameFile As String) As String
    Dim iDirPos As Integer: iDirPos = 0
    Dim strRetval As String: strRetval = ""
    Dim strPS  As String: strPS = Application.PathSeparator

    iDirPos = InStrRev(strFullNameFile, strPS)
    If iDirPos Then strRetval = Mid(strFullNameFile, 1, iDirPos)

    fnGetFilePath = strRetval

End Function


'ФУНКЦИЯ!!!! - доступна тут и у наследников
'возвращает путь до заданного вложения
Friend Function fnGetPathLevel(ByVal strFullPath As String, Optional ByVal iLvl As Integer = 9999) As String
    Dim tmpPos As Integer: tmpPos = 0
    Dim DirPos As Integer: DirPos = 0
    Dim strPS  As String: strPS = Application.PathSeparator
    Dim retval As String: retval = ""

    Do
        tmpPos = InStr(DirPos + 1, strFullPath, strPS)
        If tmpPos Then
            DirPos = tmpPos
            iLvl = iLvl - 1
            If iLvl = -1 Then Exit Do
        End If
    Loop While tmpPos
    If DirPos Then retval = Mid(strFullPath, 1, DirPos)

    fnGetPathLevel = retval

End Function
'---------------------------------------------------------
'процедура, доступна только в этом классе
'установить (дать,определить) описание ошибки
'sReplaceParam - заменяет в sErrText %1 на sReplaceParam
Private Sub sbSetError(Optional ByVal lParamErr As Long = 0, _
                       Optional ByVal sErrText As String = "", _
                       Optional ByVal sReplaceParam As String = "")
    sError = sErrText
    lErr = lParamErr
    If lParamErr <> 0 Then
        If sError <> "" Then sError = sError & vbCrLf
        sError = sError & Error(lErr) & " (" & lErr & ")"
    End If

    If InStr(sError, "%1") <> 0 Then
        sError = Replace(sError, "%1", sReplaceParam)
    End If
End Sub

'вернуть описание ошибки
Property Get GetError()
    GetError = sError
End Property

'Возвращает - True - если была ошибка
' Dim ado As New ClsADO
' ...
' ado.WorkBook = "БазаПроверок"
' If ado.IsError Then
'   тут что-то делаем в случае ошибки
' End If
Property Get IsError() As Boolean
    IsError = CBool(sError <> "" Or lErr <> 0)
End Property
'установить или отменить авто заголовки в рекордсете
'True - заголовки беруться из первой строки таблицы (диапазона)
'False - заголовки формируются и используются автоматически
'        в формате Fx - Например F1, F5, F158 ...
'возможно автоматическое переподключение
Property Let Header(ByVal bSetHeader As Boolean)
    Dim bOk    As Boolean
    If bHeader <> bSetHeader Then
        bHeader = bSetHeader
        bOk = fnReconnect()
    End If
End Property

'вернуть (показать) настройку по авто заголовкам
Property Get Header() As Boolean
    Header = bHeader
End Property
'---------------------------------------------------------
'процедура, доступна только в этом классе
'создание значений "по умолчанию"
'при создании объекта утанавливается активная рабочая книга, текущий рабочий лист
'и текщая занятая область
Sub sbCreateDefValue()
    Dim rngRaneSourse As Range

    iExcelVer = Val(Mid(Application.Version, 1, InStr(1, Application.Version, ".") - 1))
    sDataSource = ActiveWorkbook.FullName
    sSheetSource = ActiveSheet.name
    sNameStyle = "TableStyleLight10"
    bHeader = True
    bReadOnly = True
    bReplaseNull = False
    bAutoConnect = True
    lRecCount = 0
    Set rngRaneSourse = Sheets(sSheetSource).Cells(1, 1).CurrentRegion
    sUSRWhere = ""
    sUSRRange = ""
    sUSRSelect = ""
    sSelectSql = "SELECT * FROM [" & sSheetSource & "$" & Replace(rngRaneSourse.Address, "$", "") & "];"
    eLastChange = eNone
    sError = "": lErr = 0
    bRSIsReady = False
End Sub

'Отримання адреси діапазону

Function TABLEDATARANGE(ByVal loTest As ListObject) As String
'---------------------------------------------------------------------------------------------
'/// Created on:    13-Jan-2013
'/// Created by:    Zack Barresse
'/// Purpose:       Returns the range address of the specified table's data body range.
'---------------------------------------------------------------------------------------------
    If Not loTest.DataBodyRange Is Nothing Then
        TABLEDATARANGE = loTest.DataBodyRange.Address
    End If
End Function

Function TABLEHEADERRANGE(ByVal loTest As ListObject) As String
'---------------------------------------------------------------------------------------------
'/// Created on:    13-Jan-2013
'/// Created by:    Zack Barresse
'/// Purpose:       Returns the range address of the specified table's header row range.
'---------------------------------------------------------------------------------------------
    If Not loTest.HeaderRowRange Is Nothing Then
        TABLEHEADERRANGE = loTest.HeaderRowRange.Address
    End If
End Function

Function TABLETOTALSRANGE(ByVal loTest As ListObject) As String
'---------------------------------------------------------------------------------------------
'/// Created on:    13-Jan-2013
'/// Created by:    Zack Barresse
'/// Purpose:       Returns the range address of the specified table's totals row range.
'---------------------------------------------------------------------------------------------
    If Not loTest.TotalsRowRange Is Nothing Then
        TABLETOTALSRANGE = loTest.TotalsRowRange.Address
    End If
End Function
Function TABLEHEADERVISIBLE(ByVal TableCell As Range) As Boolean
'---------------------------------------------------------------------------------------------
'/// Created on:    13-Jan-2013
'/// Created by:    Zack Barresse
'/// Purpose:       Looks at a cell to see if it is part of a table and if the header is
'///                is visible.
'---------------------------------------------------------------------------------------------
    Dim loTest As ListObject

    '/// Set temporary table
    On Error Resume Next
    Set loTest = TableCell.ListObject
    On Error GoTo 0

    '/// Check if it is a table (ListObject)
    If loTest Is Nothing Then
        TABLEHEADERVISIBLE = CVErr(xlErrNA)
        Exit Function
    End If

    '/// Return visibility status of table header
    TABLEHEADERVISIBLE = loTest.ShowHeaders

End Function

Function TABLETOTALSROWVISIBLE(ByVal TableCell As Range) As Boolean
'---------------------------------------------------------------------------------------------
'/// Created on:    13-Jan-2013
'/// Created by:    Zack Barresse
'/// Purpose:       Looks at a cell to see if it is part of a table and if the totals row is
'///                is visible.
'---------------------------------------------------------------------------------------------
    Dim loTest As ListObject

    '/// Set temporary table
    On Error Resume Next
    Set loTest = TableCell.ListObject
    On Error GoTo 0

    '/// Check if it is a table (ListObject)
    If loTest Is Nothing Then
        TABLETOTALSROWVISIBLE = CVErr(xlErrNA)
        Exit Function
    End If

    '/// Return visibility status of table totalsrow
    TABLETOTALSROWVISIBLE = loTest.ShowTotals

End Function










Property Get add_PivotCache(Optional ByVal SourceType = 1, Optional SourceData) As Excel.PivotCache
'Index-number Index PivotCache
'the following line, added in a text editor,
'sets Item as the default property of the class

'SourceType
'xlConsolidation-3-Диапазоны Multiple консолидации.
'xlDatabase-1-Список Microsoft Excel или базы данных.
'xlExternal-2-Данные из другого приложения.
'xlPivotTable- -4148 Тот же источник в качестве другого отчета сводной таблицы.
'xlScenario- 4 Данные основаны на сценариях, созданных с помощью Диспетчер сценариев.
Dim pt As PivotTable
    Dim PivotSheet 'Name sheet as string
    Application.ScreenUpdating = False
'   Delete PivotSheet if it exists
    On Error Resume Next
    Application.DisplayAlerts = False
'   Set shpaste = SheetExistBookCreate(m_wb, PivotSheet, True)
    On Error GoTo 0
    


Select Case SourceType
Case 1
If IsMissing(SourceData) Then SourceData = Range("A1").CurrentRegion.Address
Set m_pc = m_wb.PivotCaches.Create( _
      SourceType:=SourceType, _
      SourceData:=SourceData)
Case 2
Set m_pc = m_wb.PivotCaches.Create( _
      SourceType:=SourceType)
PTcache.Recordset = SourceData
Case 3
Case -4148
Case 4
Case Else












'   Create a Pivot Cache
    
Set add_PivotCache = m_pc
End Property


Function CreatePivotCache(Optional SourceType = 1, Optional SourceData)

Dim PTcache As PivotCache
'SourceType
'xlConsolidation-3-Диапазоны Multiple консолидации.
'xlDatabase-1-Список Microsoft Excel или базы данных.
'xlExternal-2-Данные из другого приложения.
'xlPivotTable- -4148 Тот же источник в качестве другого отчета сводной таблицы.
'xlScenario- 4 Данные основаны на сценариях, созданных с помощью Диспетчер сценариев.
Dim pt As PivotTable
    Dim PivotSheet 'Name sheet as string
    Application.ScreenUpdating = False
'   Delete PivotSheet if it exists
    On Error Resume Next
    Application.DisplayAlerts = False
'   Set shpaste = SheetExistBookCreate(m_wb, PivotSheet, True)
    On Error GoTo 0
    


Select Case SourceType
Case 1
If IsMissing(SourceData) Then SourceData = Range("A1").CurrentRegion.Address
Set PTcache = m_wb.PivotCaches.Create( _
      SourceType:=SourceType, _
      SourceData:=SourceData)
Case 2
Set PTcache = m_wb.PivotCaches.Create( _
      SourceType:=SourceType)
PTcache.Recordset = SourceData
Case 3
Case -4148
Case 4
Case Else



End Select








'   Create a Pivot Cache
    
Set CreatePivotCache = PTcache
End Function
    
    





Sub CreatePivotTable2(PTcache As PivotCache)
    
    Dim pt As PivotTable
    Dim PivotSheet 'Name sheet as string
    Application.ScreenUpdating = False
'   Delete PivotSheet if it exists
    On Error Resume Next
    Application.DisplayAlerts = False
   Set shpaste = SheetExistBookCreate(m_wb, PivotSheet, True)
    On Error GoTo 0
    
'   Create a Pivot Cache
    Set PTcache = ActiveWorkbook.PivotCaches.Create( _
      SourceType:=xlDatabase, _
      SourceData:=Range("A1").CurrentRegion.Address)

'   Add new worksheet
    Worksheets.Add
    ActiveSheet.name = "PivotSheet"
    ActiveWindow.DisplayGridlines = False

'   Create the Pivot Table from the Cache
    Set pt = ActiveSheet.PivotTables.Add( _
      PivotCache:=PTcache, _
      TableDestination:=Range("A1"), _
      TableName:="BudgetPivot")
    
    With pt
'       Add fields
        .PivotFields("Category").Orientation = xlPageField
        .PivotFields("Division").Orientation = xlPageField
        .PivotFields("Department").Orientation = xlRowField
        .PivotFields("Month").Orientation = xlColumnField
        .PivotFields("Budget").Orientation = xlDataField
        .PivotFields("Actual").Orientation = xlDataField
        .DataPivotField.Orientation = xlRowField
    
'       Add a calculated field to compute variance
        .CalculatedFields.Add "Variance", "=Budget-Actual"
        .PivotFields("Variance").Orientation = xlDataField
        
'       Specify a number format
        .DataBodyRange.NumberFormat = "0,000"
        
'       Apply a style
        .TableStyle2 = "PivotStyleMedium2"
        
'       Hide Field Headers
        .DisplayFieldCaptions = False
        
'       Change the captions
        .PivotFields("Sum of Budget").Caption = " Budget"
        .PivotFields("Sum of Actual").Caption = " Actual"
        .PivotFields("Sum of Variance").Caption = " Variance"
    End With
End Sub



Function Add_PivotTable(pvtCache As PivotCache, TableName, Optional TableDestination)
'PURPOSE: Creates a brand new Pivot table on a new worksheet from data in the ActiveSheet
'Source: www.TheSpreadsheetGuru.com

Dim sht As Worksheet
'Dim pvtCache As PivotCache
Dim pvt As PivotTable
Dim StartPvt As String
Dim SrcData As String

'Determine the data range you want to pivot
'  SrcData = ActiveSheet.Name & "!" & Range("A1:R100").Address(ReferenceStyle:=xlR1C1)

'Create a new worksheet
Set sht = SheetExistBookCreate(m_wb, TableName, True)

'Where do you want Pivot Table to start?
  StartPvt = sht.name & "!" & sht.Range("A3").Address(ReferenceStyle:=xlR1C1)

'Create Pivot Cache from Source Data
'  Set pvtCache = ActiveWorkbook.PivotCaches.Create( _
'    SourceType:=xlDatabase, _
'    SourceData:=SrcData)

'Create Pivot table from Pivot Cache
  Set pvt = pvtCache.CreatePivotTable( _
    TableDestination:=StartPvt, _
    TableName:="CV_" & TableName)

End Function



'Example Data in Sheet1
'Department  Region     Profit
'109            5         119000
'107            3         64000
'109            2         12000
Sub CreatePivotsb()
'Declaration
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
'Adding new worksheet
Set ws = Worksheets.Add
'Creating Pivot cache
Set pc = ActiveWorkbook.PivotCaches.Create(xlDatabase, "Sheet1!R1C1:R10C3")
'Creating Pivot table
Set pt = pc.CreatePivotTable(ws.Range("B3"))
'Setting Fields
With pt
'set row field
With .PivotFields("Department")
.Orientation = xlRowField
.Position = 1
End With
'set column field
With .PivotFields("Region")
.Orientation = xlColumnField
.Position = 1
End With
'set data field
.AddDataField .PivotFields("Profit"), "Sum of Profit", xlSum
End With
End Sub
Function ShowCacheIndex(rngPT As Range) As Long
'Показать CacheIndex сводной таблице
'0 если rngPT не пренадлежит любой сводной таблице
'= ShowCacheIndex (А3)

On Error Resume Next
  ShowCacheIndex = rngPT.PivotTable.CacheIndex
If Err <> 0 Then ShowCacheIndex = 0
End Function
'Показать память Pivot кэш, используемый
Function GetMemory(rngPT As Range) As Long
'pivot table tutorial by contextures.com
'Показать память Pivot кэш, используемый
'Результат отображается в килобайтах /1000
'= GetMemory (А3) / 1000
  Dim pt As PivotTable
  Set pt = rngPT.PivotTable
  GetMemory = ActiveWorkbook _
    .PivotCaches(pt.CacheIndex).MemoryUsed
End Function

'Показать граф Pivot кэша
'Sub CountCaches()
'  MsgBox "There are " _
'      & ActiveWorkbook.PivotCaches.Count _
'      & " pivot caches in the active workook."
'End Sub

Function GetRecords(rngPT As Range) As Long
'Можно отобразить количество записей в кэше сводной,
'используя следующие пользовательские функции.
'= GetRecords (А3)

'pivot table tutorial by contextures.com
  Dim pt As PivotTable
  Set pt = rngPT.PivotTable
  GetRecords = ActiveWorkbook _
    .PivotCaches(pt.CacheIndex).RecordCount
End Function

Sub ChangePivotCache()
'Если вы создали несколько сводных таблиц в книге,
'вы можете найти его более эффективно использовать тот же кэш поворота для всех сводных таблиц.
'Следующий код изменит кэш поворота для каждой сводной таблицы в книге

'pivot table tutorial by contextures.com
'change pivot cache for all Pivot Tables in workbook
Dim pt As PivotTable
Dim wks As Worksheet

  For Each wks In ActiveWorkbook.Worksheets
    For Each pt In wks.PivotTables
        pt.CacheIndex = Sheets("Pivot").PivotTables(1).CacheIndex
    Next pt
  Next wks

End Sub
Sub SelPTNewCache()
'Создать новый Pivot кэш для выбранной сводной таблицы
'Если два или более сводных таблиц основаны на одном сводном кэше,
'они будут делить некоторые функции,
'такие как вычисляемых элементов и сгруппированных полей.
'
'Чтобы создать отдельный кэш поворота для сводной таблицы,
'вы можете выбрать ячейку в сводной таблице, а затем запустить следующий код.
'
'Код добавляет лист к книге, и создает новую сводную таблицу там,
'основанную на том же источнике данных, но в новом сводном кэше.
'
'Выбранная сводной таблицы устанавливается в том же сводном кэше
'в качестве новой таблицы, а затем временный лист будет удален.









    Dim wsTemp As Worksheet
    Dim pt As PivotTable
    
    On Error Resume Next
    Set pt = ActiveCell.PivotTable
    
    If pt Is Nothing Then
        MsgBox "Active cell is not in a pivot table"
    Else
        Set wsTemp = Worksheets.Add
        
        ActiveWorkbook.PivotCaches.Create( _
            SourceType:=xlDatabase, _
            SourceData:=pt.SourceData).CreatePivotTable _
            TableDestination:=wsTemp.Range("A3"), _
            TableName:="PivotTableTemp"
        
        pt.CacheIndex = wsTemp.PivotTables(1).CacheIndex
        
        Application.DisplayAlerts = False
        wsTemp.Delete
        Application.DisplayAlerts = True
    End If
    
exitHandler:
        Set pt = Nothing

End Sub

Sub CheckCaches()
'Удалить дубликаты Pivot кэши
'Несколько сводных таблиц в рабочей книге может быть основан на том же источнике данных,
'но используют различные кэши шарнирные.
'Этот макрос создает список шарнирных кэшей,
'чеки для дублированных источников данных, а также устраняет дубликаты кэшей.
' Developed by Contextures Inc.
' www.contextures.com
Dim pc As PivotCache
Dim wsList As Worksheet
Dim lRow As Long
Dim lRowPC As Long
Dim pt As PivotTable
Dim ws As Worksheet
Dim lStart As Long
lStart = 2
lRow = lStart

Set wsList = Worksheets.Add
For Each pc In ActiveWorkbook.PivotCaches
  wsList.Cells(lRow, 1).value = pc.Index
  wsList.Cells(lRow, 2).value = pc.SourceData
  wsList.Cells(lRow, 3).FormulaR1C1 = _
    "=INDEX(R1C[-2]:R[-1]C[-2],MATCH(RC[-1],R1C[-1]:R[-1]C[-1],0))"
  lRow = lRow + 1
Next pc

For lRowPC = lRow - 1 To lStart Step -1
  With wsList.Cells(lRowPC, 3)
    If IsNumeric(.value) Then
      For Each ws In ActiveWorkbook.Worksheets
      Debug.Print ws.name
        For Each pt In ws.PivotTables
        Debug.Print .Offset(0, -2).value
          If pt.CacheIndex = .Offset(0, -2).value Then
            pt.CacheIndex = .value
          End If
        Next pt
      Next ws
    End If
  End With
Next lRowPC

'uncomment lines below to delete the temp worksheet
'Application.DisplayAlerts = False
'wsList.Delete

exitHandler:
Application.DisplayAlerts = True
Exit Sub

errHandler:
MsgBox "Could not change all pivot caches"
Resume exitHandler

End Sub
Sub DeleteOldItemsWB()
'Программным Очистить старые элементы - Excel 2002 или более поздняя версия

'pivot table tutorial by contextures.com
'gets rid of unused items in PivotTable
' based on MSKB (202232)
Dim ws As Worksheet
Dim pt As PivotTable
Dim pf As PivotField
Dim pi As PivotItem

On Error Resume Next
For Each ws In ActiveWorkbook.Worksheets
  For Each pt In ws.PivotTables
    pt.RefreshTable
    pt.ManualUpdate = True
    For Each pf In pt.VisibleFields
      If pf.name <> "Data" Then
        For Each pi In pf.PivotItems
          If pi.RecordCount = 0 And _
            Not pi.IsCalculated Then
            pi.Delete
          End If
        Next pi
      End If
    Next pf
    pt.ManualUpdate = False
    'pt.RefreshTable 'optional - might hang Excel
                 'if 2 or more pivot tables on one sheet
  Next pt
Next ws

End Sub
Sub DeleteAllPivotTables()
'PURPOSE: Delete all Pivot Tables in your Workbook
'SOURCE: www.TheSpreadsheetGuru.com

Dim sht As Worksheet
Dim pvt As PivotTable

'Loop Through Each Pivot Table In Currently Viewed Workbook
  For Each sht In ActiveWorkbook.Worksheets
    For Each pvt In sht.PivotTables
      pvt.TableRange2.Clear
    Next pvt
  Next sht
  
End Sub
Sub DeletePivotTable()
'PURPOSE: How to delete a specifc Pivot Table
'SOURCE: www.TheSpreadsheetGuru.com

'Delete Pivot Table By Name
  ActiveSheet.PivotTables("PivotTable1").TableRange2.Clear

End Sub
Sub Adding_PivotFields()
'PURPOSE: Show how to add various Pivot Fields to Pivot Table
'SOURCE: www.TheSpreadsheetGuru.com

Dim pvt As PivotTable

Set pvt = ActiveSheet.PivotTables("PivotTable1")
    
  'Add item to the Report Filter
    pvt.PivotFields("Year").Orientation = xlPageField
  
  'Add item to the Column Labels
    pvt.PivotFields("Month").Orientation = xlColumnField
    
  'Add item to the Row Labels
    pvt.PivotFields("Account").Orientation = xlRowField
    
  'Position Item in list
    pvt.PivotFields("Year").Position = 1
    
  'Format Pivot Field
    pvt.PivotFields("Year").NumberFormat = "#,##0"
    
  'Turn on Automatic updates/calculations --like screenupdating to speed up code
    pvt.ManualUpdate = False
    
End Sub
Sub AddCalculatedField()
'PURPOSE: Add a calculated field to a pivot table
'SOURCE: www.TheSpreadsheetGuru.com
   
Dim pvt As PivotTable
Dim pf As PivotField

'Set Variable to Desired Pivot Table
  Set pvt = ActiveSheet.PivotTables("PivotTable1")

'Set Variable Equal to Desired Calculated Pivot Field
  For Each pf In pvt.PivotFields
    If pf.SourceName = "Inflation" Then Exit For
  Next

'Add Calculated Field to Pivot Table
  pvt.AddDataField pf

End Sub
Sub AddValuesField()
'PURPOSE: Add A Values Field to a Pivot Table
'SOURCE: www.TheSpreadsheetGuru.com

Dim pvt As PivotTable
Dim pf As String
Dim pf_Name As String

pf = "Salaries"
pf_Name = "Sum of Salaries"

Set pvt = ActiveSheet.PivotTables("PivotTable1")

pvt.AddDataField pvt.PivotFields("Salaries"), pf_Name, xlSum

End Sub
Sub RemovePivotField()
'PURPOSE: Remove a field from a Pivot Table
'SOURCE: www.TheSpreadsheetGuru.com

'Removing Filter, Columns, Rows
  ActiveSheet.PivotTables("PivotTable1").PivotFields("Year").Orientation = xlHidden
    
'Removing Values
  ActiveSheet.PivotTables("PivotTable1").PivotFields("Sum of Salaries").Orientation = xlHidden
  
End Sub
Sub RemoveCalculatedField()
'PURPOSE: Remove a calculated field from a pivot table
'SOURCE: www.TheSpreadsheetGuru.com

Dim pvt As PivotTable
Dim pf As PivotField
Dim pi As PivotItem

'Set Variable to Desired Pivot Table
  Set pvt = ActiveSheet.PivotTables("PivotTable1")

'Set Variable Equal to Desired Calculated Data Field
  For Each pf In pvt.DataFields
    If pf.SourceName = "Inflation" Then Exit For
  Next

'Hide/Remove the Calculated Field
  pf.DataRange.Cells(1, 1).PivotItem.Visible = False

End Sub
Sub ReportFiltering_Single()
'PURPOSE: Filter on a single item with the Report Filter field
'SOURCE: www.TheSpreadsheetGuru.com

Dim pf As PivotField

Set pf = ActiveSheet.PivotTables("PivotTable2").PivotFields("Fiscal_Year")

'Clear Out Any Previous Filtering
  pf.ClearAllFilters

'Filter on 2014 items
  pf.CurrentPage = "2014"

End Sub
Sub ReportFiltering_Multiple()
'PURPOSE: Filter on multiple items with the Report Filter field
'SOURCE: www.TheSpreadsheetGuru.com

Dim pf As PivotField

Set pf = ActiveSheet.PivotTables("PivotTable2").PivotFields("Variance_Level_1")

'Clear Out Any Previous Filtering
  pf.ClearAllFilters

'Enable filtering on multiple items
    pf.EnableMultiplePageItems = True
    
'Must turn off items you do not want showing
    pf.PivotItems("Jan").Visible = False
    pf.PivotItems("Feb").Visible = False
    pf.PivotItems("Mar").Visible = False

End Sub
Sub ClearReportFiltering()
'PURPOSE: How to clear the Report Filter field
'SOURCE: www.TheSpreadsheetGuru.com

Dim pf As PivotField

Set pf = ActiveSheet.PivotTables("PivotTable2").PivotFields("Fiscal_Year")

'Option 1: Clear Out Any Previous Filtering
  pf.ClearAllFilters
  
'Option 2: Show All (remove filtering)
  pf.CurrentPage = "(All)"

End Sub
Sub RefreshingPivotTables()
'PURPOSE: Shows various ways to refresh Pivot Table Data
'SOURCE: www.TheSpreadsheetGuru.com

'Refresh A Single Pivot Table
  ActiveSheet.PivotTables("PivotTable1").PivotCache.Refresh
  
'Refresh All Pivot Tables
  ActiveWorkbook.RefreshAll
    
End Sub
Sub ChangePivotDataSourceRange()
'PURPOSE: Change the range a Pivot Table pulls from
'SOURCE: www.TheSpreadsheetGuru.com

Dim sht As Worksheet
Dim SrcData As String
Dim pvtCache As PivotCache

'Determine the data range you want to pivot
  Set sht = ThisWorkbook.Worksheets("Sheet1")
  SrcData = sht.name & "!" & Range("A1:R100").Address(ReferenceStyle:=xlR1C1)
  
'Create New Pivot Cache from Source Data
  Set pvtCache = ActiveWorkbook.PivotCaches.Create( _
    SourceType:=xlDatabase, _
    SourceData:=SrcData)

'Change which Pivot Cache the Pivot Table is referring to
  ActiveSheet.PivotTables("PivotTable1").ChangePivotCache (pvtCache)

End Sub

Sub PivotGrandTotals(Index As Variant)
'PURPOSE: Show setup for various Pivot Table Grand Total options
'SOURCE: www.TheSpreadsheetGuru.com

Dim pvt As PivotTable

Set pvt = m_Tables(Index)

'Off for Rows and Columns
  pvt.ColumnGrand = False
  pvt.RowGrand = False

'On for Rows and Columns
  pvt.ColumnGrand = True
  pvt.RowGrand = True

'On for Rows only
  pvt.ColumnGrand = False
  pvt.RowGrand = True
  
'On for Columns Only
  pvt.ColumnGrand = True
  pvt.RowGrand = False

End Sub
Sub PivotReportLayout(Index As Variant)
'PURPOSE: Show setup for various Pivot Table Report Layout options
'SOURCE: www.TheSpreadsheetGuru.com

Dim pvt As PivotTable

Set pvt = m_Tables(Index)

'Show in Compact Form
  pvt.RowAxisLayout xlCompactRow

'Show in Outline Form
  pvt.RowAxisLayout xlOutlineRow
  
'Show in Tabular Form
  pvt.RowAxisLayout xlTabularRow
    
End Sub
Sub PivotTable_DataFormatting(Index As Variant)
'PURPOSE: Various ways to format a Pivot Table's data
'SOURCE: www.TheSpreadsheetGuru.com

Dim pvt As PivotTable

Set pvt = m_Tables(Index)

'Change Data's Number Format
  pvt.DataBodyRange.NumberFormat = "#,##0;(#,##0)"

'Change Data's Fill Color
  pvt.DataBodyRange.Interior.Color = RGB(0, 0, 0)

'Change Data's Font Type
  pvt.DataBodyRange.Font.FontStyle = "Arial"

End Sub
Sub PivotField_DataFormatting(Index As Variant, PivotFieldName)
'PURPOSE: Various ways to format a Pivot Field's data
'SOURCE: www.TheSpreadsheetGuru.com

Dim pf As PivotField

Set pf = m_Tables(Index).PivotFields(PivotFieldName)

'Change Data's Number Format
  pf.DataRange.NumberFormat = "#,##0;(#,##0)"
  
'Change Data's Fill Color
  pf.DataRange.Interior.Color = RGB(219, 229, 241)

'Change Data's Font Type
  pf.DataRange.Font.FontStyle = "Arial"

End Sub
Sub PivotField_ExpandCollapse(Index As Variant, PivotFieldName)
'PURPOSE: Shows how to Expand or Collapse the detail of a Pivot Field
'SOURCE: www.TheSpreadsheetGuru.com

Dim pf As PivotField

Set pf = m_Tables(Index).PivotFields(PivotFieldName)

'Collapse Pivot Field
  pf.ShowDetail = False
  
'Expand Pivot Field
  pf.ShowDetail = True

End Sub
Sub RepeatLabels(Index As Variant)
'
' Макрос2 Макрос
'

'
    m_Tables(Index).RepeatAllLabels xlRepeatLabels
End Sub
Sub CurrentPageSelect(Index As Variant, CurrentPageName, Optional CurrentPageValue = "(All)", Optional bEnableMultiplePageItems = False, Optional bRepeatLabels = True)
Dim twb As Workbook
Dim tableSheet As Worksheet
Set twb = ThisWorkbook
Dim ptj As clsmPivotTables
Dim pc As PivotCache
Dim pt As PivotTable
Dim pts As PivotTables
Dim pf As PivotField
Dim pi As PivotItem
Dim pis As PivotItem
Dim obj As Object
Set r = Cells.Columns(1)
Set ptj = New clsmPivotTables
With ptj
.Initialize twb
'Debug.Print "Count PivotCaches-" & .Count_PivotCaches
Debug.Print PivotTableName & "-" & .Exists(PivotTableName)
Set pt = .Items(Index)
Set tableSheet = pt.Parent
Set pf = pt.PageFields(CurrentPageName)
pf.EnableMultiplePageItems = False
pf.CurrentPage = CurrentPageValue
pf.EnableMultiplePageItems = bEnableMultiplePageItems
If bRepeatLabels Then RepeatLabels (Index)
End With
End Sub



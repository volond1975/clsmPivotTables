'#Modifi
'25.08.2016
'add Function LogikVBA


Function LogikVBA(v) As Boolean
'Возвращает Истину или ложь в зависмости от значения v
Select Case v
Case ""
LogikVBA = False
Case 0

LogikVBA = False
Case 1
LogikVBA = True
Case "Yes"
LogikVBA = True
Case "No"
LogikVBA = False
Case "Да"
LogikVBA = True
Case "Нет"
LogikVBA = False
Case "Так"
LogikVBA = True
Case "Ні"
LogikVBA = False


Case Else
LogikVBA = False
End Select



End Function



'работает аналогично ВПР, но возвращает массив данных
Function VLOOKUP3(Table As Range, SearchColumnNum As Integer, _
SearchValue As Variant, ResultColumnNum As Integer)
    
    Dim i, j As Integer
    Dim out(1000) As Variant
    Dim rCol As Range
    j = 0
        For i = 1 To Table.Rows.Count
            If Table.Cells(i, SearchColumnNum) = SearchValue Then
                out(j) = Table.Cells(i, ResultColumnNum)
                j = j + 1
            End If
        Next i
    VLOOKUP3 = Application.Transpose(out)
End Function

'Транслитерация русского текста в английский
Function Translit(txt As String) As String
    Dim Rus As Variant
    Rus = Array("а", "б", "в", "г", "д", "е", "ё", "ж", "з", "и", "й", "к", "л", "м", "н", "о", "п", "р", "с", "т", "у", "ф", "х", "ц", "ч", "ш", "щ", "ъ", "ы", "ь", "э", "ю", "я", "А", "Б", "В", "Г", "Д", "Е", "Ё", "Ж", "З", "И", "Й", "К", "Л", "М", "Н", "О", "П", "Р", "С", "Т", "У", "Ф", "Х", "Ц", "Ч", "Ш", "Щ", "Ъ", "Ы", "Ь", "Э", "Ю", "Я")
    Dim Eng As Variant
    Eng = Array("a", "b", "v", "g", "d", "e", "jo", "zh", "z", "i", "j", "k", "l", "m", "n", "o", "p", "r", "s", "t", "u", "f", "kh", "ts", "ch", "sh", "sch", "''", "y", "'", "e", "ju", "ja", "A", "B", "V", "G", "D", "E", "JO", "ZH", "Z", "I", "J", "K", "L", "M", "N", "O", "P", "R", "S", "T", "U", "F", "KH", "TS", "CH", "SH", "SCH", "''", "Y", "'", "E", "JU", "JA")
    
    For i = 1 To Len(txt)
        с = Mid(txt, i, 1)
    
        flag = 0
        For j = 0 To 64
            If Rus(j) = с Then
                outchr = Eng(j)
                flag = 1
                Exit For
            End If
        Next j
        If flag Then outstr = outstr & outchr Else outstr = outstr & с
    Next i
    
    Translit = outstr
    
End Function

'слияние текста всех ячеек диапазона с разделителем
Function MultiCat(ByRef rng As Excel.Range, Optional ByVal DELIM As String = "") As String
     Dim rcell As Range
     For Each rcell In rng
         MultiCat = MultiCat & DELIM & rcell.Text
     Next rcell
     MultiCat = Mid(MultiCat, Len(DELIM) + 1)
  End Function

'вывод заднного количества неповторяющихся случайных чисел из диапазона
Function Lotto(Bottom As Integer, Top As Integer, Amount As Integer)
    Dim iArr As Variant
    Dim i As Integer
    Dim r As Integer
    Dim temp As Integer
    Dim out(1000) As Variant
    
    Application.Volatile
    
    ReDim iArr(Bottom To Top)
    For i = Bottom To Top
        iArr(i) = i
    Next i
    
    For i = Top To Bottom + 1 Step -1
        r = Int(Rnd() * (i - Bottom + 1)) + Bottom
        temp = iArr(r)
        iArr(r) = iArr(i)
        iArr(i) = temp
    Next i
    j = 0
    For i = Bottom To Bottom + Amount - 1
        out(j) = iArr(i)
        j = j + 1
    Next i
    
    Lotto = Application.Transpose(out)
    
End Function
'выбор случайного элемента из диапазона
Function RandomSelect(TargetCells)
    RandomSelect = TargetCells.Cells(Int(Rnd * TargetCells.Count) + 1)
End Function

'вывод дня недели по дате словом
Function WeekdayWord(MyDate As Date) As String
    Dim days As Variant
    days = Array("понедельник", "вторник", "среда", "четверг", "пятница", "суббота", "воскресенье")
    WeekdayWord = days(Weekday(MyDate, vbMonday) - 1)
End Function

'выводит любой заданный разряд числа
Function Class(m, i)
       Class = Int(Int(m - (10 ^ i) * Int(m / (10 ^ i))) / 10 ^ (i - 1))
End Function

'сумма прописью на русском языке
Function PropisRus(n As Double, rub As Boolean) As String
    Dim Nums1, Nums2, Nums3, Nums4 As Variant
    Nums1 = Array("", "один ", "два ", "три ", "четыре ", "пять ", "шесть ", "семь ", "восемь ", "девять ")
    Nums2 = Array("", "десять ", "двадцать ", "тридцать ", "сорок ", "пятьдесят ", "шестьдесят ", "семьдесят ", "восемьдесят ", "девяносто ")
    Nums3 = Array("", "сто ", "двести ", "триста ", "четыреста ", "пятьсот ", "шестьсот ", "семьсот ", "восемьсот ", "девятьсот ")
    Nums4 = Array("", "одна ", "две ", "три ", "четыре ", "пять ", "шесть ", "семь ", "восемь ", "девять ")
    Nums5 = Array("десять ", "одиннадцать ", "двенадцать ", "тринадцать ", "четырнадцать ", "пятнадцать ", "шестнадцать ", "семнадцать ", "восемнадцать ", "девятнадцать ")
    
    If n <= 0 Then
        Propis = "ноль"
        Exit Function
    End If
    ed = Class(n, 1)
    dec = Class(n, 2)
    sot = Class(n, 3)
    tys = Class(n, 4)
    dectys = Class(n, 5)
    sottys = Class(n, 6)
    mil = Class(n, 7)
    decmil = Class(n, 8)
    
    Select Case decmil
        Case 1
            mil_txt = Nums5(mil) & "миллионов "
            GoTo www
        Case 2 To 9
            decmil_txt = Nums2(decmil)
    End Select
    
    Select Case mil
        Case 1
            mil_txt = Nums1(mil) & "миллион "
        Case 2, 3, 4
            mil_txt = Nums1(mil) & "миллиона "
        Case 5 To 20
            mil_txt = Nums1(mil) & "миллионов "
    End Select
www:
    sottys_txt = Nums3(sottys)
    Select Case dectys
        Case 1
            tys_txt = Nums5(tys) & "тысяч "
            GoTo eee
        Case 2 To 9
            dectys_txt = Nums2(dectys)
    End Select
    
    Select Case tys
        Case 0
            If dectys > 0 Then tys_txt = Nums4(tys) & "тысяч "
        Case 1
            tys_txt = Nums4(tys) & "тысячa "
        Case 2, 3, 4
            tys_txt = Nums4(tys) & "тысячи "
        Case 5 To 9
            tys_txt = Nums4(tys) & "тысяч "
    End Select
    If dectys = 0 And tys = 0 And sottys <> 0 Then sottys_txt = sottys_txt & " тысяч "
eee:
    sot_txt = Nums3(sot)
    
    Select Case dec
    Case 1
        ed_txt = Nums5(ed)
        GoTo rrr
    Case 2 To 9
        dec_txt = Nums2(dec)
    End Select
    
    ed_txt = Nums1(ed)
rrr:
    If rub Then
        Select Case ed_txt
            Case "один "
                rub_txt = "рубль"
            Case "два ", "три ", "четыре "
                rub_txt = "рубля"
            Case Else
                rub_txt = "рублей"
        End Select
        kops = Round((n * 100 - Int(n) * 100), 0)
        PropisRus = decmil_txt & mil_txt & sottys_txt & dectys_txt & tys_txt & sot_txt & dec_txt & ed_txt & rub_txt & " " & kops & " коп."
    Else
        PropisRus = decmil_txt & mil_txt & sottys_txt & dectys_txt & tys_txt & sot_txt & dec_txt & ed_txt
    End If
End Function

'сумма прописью на английском языке
Function PropisEng(ByVal strAmount As String, strCur As String, strDec As String, iPrec As Integer)
    Dim BigDenom As String, SmallDenom As String, temp As String
    Dim iDecimalPlace As Integer
    Dim Count As Integer
    
    ReDim Place(9) As String
    Place(2) = " Thousand "
    Place(3) = " Million "
    Place(4) = " Billion "
    Place(5) = " Trillion "
    
    ' String representation of amount.
    strAmount = Trim(str(strAmount))
    
    ' Position of decimal place 0 if none.
    iDecimalPlace = InStr(strAmount, ".")
    
    ' Separate the Integer part from the decimals.
    If iDecimalPlace > 0 Then
        SmallDenom = Left(Right(strAmount, Len(strAmount) - iDecimalPlace) & "0000000000", iPrec)
        SmallDenom = PropisEng(SmallDenom, strDec, "", 0)
        BigDenom = Left(strAmount, iDecimalPlace - 1)
        BigDenom = PropisEng(BigDenom, strCur, "", 0)
        PropisEng = BigDenom & " And " & SmallDenom
        Exit Function
    End If
    If iDecimalPlace = 0 Then
        Count = 1
        Do While strAmount <> ""
            temp = GetHundreds(Right(strAmount, 3))
            If temp <> "" Then BigDenom = temp & Place(Count) & BigDenom
            If Len(strAmount) > 3 Then
                strAmount = Left(strAmount, Len(strAmount) - 3)
            Else
                strAmount = ""
            End If
            Count = Count + 1
        Loop
        Select Case BigDenom
            Case ""
                BigDenom = "No " & strCur
            Case "One"
                BigDenom = "One " & strCur
             Case Else
                BigDenom = BigDenom & " " & strCur
        End Select
        PropisEng = BigDenom
    End If
End Function

' Converts a number from 100-999 into text
Function GetHundreds(ByVal MyNumber)
    Dim result As String
    If Val(MyNumber) = 0 Then Exit Function
    MyNumber = Right("000" & MyNumber, 3)
    ' Convert the hundreds place.
    If Mid(MyNumber, 1, 1) <> "0" Then
        result = GetDigit(Mid(MyNumber, 1, 1)) & " Hundred "
    End If
    ' Convert the tens and ones place.
    If Mid(MyNumber, 2, 1) <> "0" Then
        result = result & GetTens(Mid(MyNumber, 2))
    Else
        result = result & GetDigit(Mid(MyNumber, 3))
    End If
    GetHundreds = result
End Function

' Converts a number from 10 to 99 into text.
Function GetTens(TensText)
    Dim result As String
    result = ""           ' Null out the temporary function value."
    If Val(Left(TensText, 1)) = 1 Then   ' If value between 10-19…
        Select Case Val(TensText)
            Case 10: result = "Ten"
            Case 11: result = "Eleven"
            Case 12: result = "Twelve"
            Case 13: result = "Thirteen"
            Case 14: result = "Fourteen"
            Case 15: result = "Fifteen"
            Case 16: result = "Sixteen"
            Case 17: result = "Seventeen"
            Case 18: result = "Eighteen"
            Case 19: result = "Nineteen"
            Case Else
        End Select
    Else                                 ' If value between 20-99…
        Select Case Val(Left(TensText, 1))
            Case 2: result = "Twenty "
            Case 3: result = "Thirty "
            Case 4: result = "Forty "
            Case 5: result = "Fifty "
            Case 6: result = "Sixty "
            Case 7: result = "Seventy "
            Case 8: result = "Eighty "
            Case 9: result = "Ninety "
            Case Else
        End Select
        result = result & GetDigit _
            (Right(TensText, 1))  ' Retrieve ones place.
    End If
    GetTens = result
End Function

' Converts a number from 1 to 9 into text.
Function GetDigit(Digit)
    Select Case Val(Digit)
        Case 1: GetDigit = "One"
        Case 2: GetDigit = "Two"
        Case 3: GetDigit = "Three"
        Case 4: GetDigit = "Four"
        Case 5: GetDigit = "Five"
        Case 6: GetDigit = "Six"
        Case 7: GetDigit = "Seven"
        Case 8: GetDigit = "Eight"
        Case 9: GetDigit = "Nine"
        Case Else: GetDigit = ""
    End Select
End Function

'суммирует ячейки в заданном интервале
Function SumBetween(TargetCells As Range, min As Long, max As Long, IncludeMin As Boolean, IncludeMax As Boolean) As Long
    Dim s As Long
    For Each c In TargetCells
        If IncludeMin And IncludeMax = True Then If c >= min And c <= max Then s = s + c
        If IncludeMin And Not IncludeMax Then If c >= min And c < max Then s = s + c
        If Not IncludeMin And IncludeMax Then If c > min And c <= max Then s = s + c
        If Not IncludeMin And Not IncludeMax Then If c > min And c < max Then s = s + c
    Next c
    SumBetween = s
End Function

'возвращает дату N-го дня недели (W) для заданного месяца М и года Y
Function NeedDate(n As Integer, w As Integer, m As Integer, Y As Integer) As Date
    Dim i, md As Integer
    Dim d As Date
    'определяем сколько дней в месяце
    Select Case m
    Case 1, 3, 5, 7, 8, 10, 12
        md = 31
    Case 4, 6, 9, 11
        md = 30
    Case 2
        'если год високосный, то в феврале 29 иначе 28 дней
        If (Y - 2000) Mod 4 = 0 Then md = 29 Else md = 28
    End Select
    
    For d = DateSerial(Y, m, 1) To DateSerial(Y, m, md)
        If Weekday(d, vbMonday) = w Then
            i = i + 1
            If i = n Then
                NeedDate = d
                Exit Function
            End If
        End If
    Next d
    NeedDate = " "
End Function

'выделяет числа из ячейки
Function GetNumbers(TargetCell As Range) As String
    Dim LenStr As Long
    For LenStr = 1 To Len(TargetCell)
        Select Case Asc(Mid(TargetCell, LenStr, 1))
        Case 48 To 57
            GetNumbers = GetNumbers & Mid(TargetCell, LenStr, 1)
        End Select
    Next
End Function

'выделяет текст из ячейки
Function GetText(TargetCell As Range) As String
    Dim LenStr As Long
    For LenStr = 1 To Len(TargetCell)
        Select Case Asc(Mid(TargetCell, LenStr, 1))
        Case 65 To 90
            GetText = GetText & Mid(TargetCell, LenStr, 1)
        Case 97 To 122
            GetText = GetText & Mid(TargetCell, LenStr, 1)
        Case 192 To 255
            GetText = GetText & Mid(TargetCell, LenStr, 1)
        End Select
    Next
End Function

'возвращает первое значение в указанной строке
Function FirstInRow(myRow As Range)
    If Cells(myRow.Row, 1) <> "" Then FirstInRow = Cells(myRow.Row, 1).value
    If Cells(myRow.Row, 1) = "" Then FirstInRow = Cells(myRow.Row, 1).End(xlToRight).value
End Function

'возвращает первое значение в указанном столбце
Function FirstInColumn(myColumn As Range)
    If Cells(1, myColumn.Column) <> "" Then FirstInColumn = Cells(1, myColumn.Column).value
    If Cells(1, myColumn.Column) = "" Then FirstInColumn = Cells(1, myColumn.Column).End(xlDown).value
End Function

'возвращает последнее значение в указанной строке
Function LastInRow(myRow As Range)
    If Cells(myRow.Row, Sheets(1).Columns.Count) <> "" Then LastInRow = Cells(myRow.Row, Sheets(1).Columns.Count).value
    If Cells(myRow.Row, Sheets(1).Columns.Count) = "" Then LastInRow = Cells(myRow.Row, Sheets(1).Columns.Count).End(xlToLeft).value
End Function

'возвращает последнее значение в указанном столбце
Function LastInColumn(myColumn As Range)
    If Cells(Sheets(1).Rows.Count, myColumn.Column) <> "" Then LastInColumn = Cells(Sheets(1).Rows.Count, myColumn.Column).value
    If Cells(Sheets(1).Rows.Count, myColumn.Column) = "" Then LastInColumn = Cells(Sheets(1).Rows.Count, myColumn.Column).End(xlUp).value
End Function

'возвращает имя листа
Function SheetName1() As String
    SheetName = ActiveSheet.name
End Function

'возвращает имя книги
Function WorkbookName() As String
    WorkbookName = ActiveWorkbook.name
End Function

'возвращает полное имя файла (полный путь)
Function FullFileName() As String
    FullFileName = ActiveWorkbook.FullName
End Function

'возвращает имя текущего пользователя
Function UserName() As String
    UserName = Application.UserName
End Function

'код цвета заливки ячейки
Function CellColor(cell As Range)
    CellColor = cell.Interior.ColorIndex
End Function

'код цвета шрифта ячейки
Function CellFontColor(cell As Range)
    CellFontColor = cell.Font.ColorIndex
End Function


'выводит текущие условия автофильтра
Function AutoFilter_Criteria(Header As Range) As String
Dim strCri1 As String, strCri2 As String
    Application.Volatile
    With Header.Parent.AutoFilter
        With .Filters(Header.Column - .Range.Column + 1)
            If Not .On Then Exit Function
                strCri1 = .Criteria1
            If .Operator = xlAnd Then
                strCri2 = " AND " & .Criteria2
            ElseIf .Operator = xlOr Then
                strCri2 = " OR " & .Criteria2
            End If
        End With
    End With
    AutoFilter_Criteria = UCase(Header) & ": " & strCri1 & strCri2
End Function

'выделяет подстроку из строки
Public Function Substring(txt, Delimiter, n) As String
Dim x As Variant
    x = Split(txt, Delimiter)
    If n > 0 And n - 1 <= UBound(x) Then
        Substring = x(n - 1)
    Else
        Substring = ""
    End If
End Function
Public Function GrupString() As String
Dim x As Variant
Dim r As Variant
Dim v() As Variant
Dim vv() As Variant
Dim R_count As Long
Dim R_R As Range
Dim max_R As Integer
Set r = Selection
R_count = r.Count
ReDim Preserve vv(R_count - 1)
k = 0
max_R = 0
For Each R_R In r.Cells
z = Split(R_R.value, " ")
   vv(k) = UBound(z) + 1
k = k + 1
    Next
    maxvv = Application.WorksheetFunction.max(vv)
ReDim Preserve v(R_count - 1, maxvv - 1)
k = 0
For Each R_R In r.Cells
z = Split(R_R.value, " ")
For j = 0 To UBound(z)
   v(k, j) = z(j)
   Next j
k = k + 1
    Next
  k = 0
  w = 0
  For j = 0 To maxvv - 1
 k = 0
For Each R_R In r.Cells
zn = v(0, j)

  If v(k, j) <> zn Then
  w = j
  n = 1
  Exit For
  End If
  
   
k = k + 1
    Next
   If w <> 0 Then
   Exit For
   
  End If
   Next j
   If n = 0 Then
   w = maxvv
   End If
   If w <> 0 Then
   For i = 0 To w - 1
   Text = Text & " " & v(0, i)
   GrupString = Trim(Text)
   Next i
   Else
   GrupString = ""
   End If

End Function




'усовершенствованная версия ВПР
Function VLOOKUP2(Table As Range, SearchColumnNum As Integer, SearchValue As Variant, n As Integer, ResultColumnNum As Integer)

    Dim i As Integer
    Dim iCount As Integer
    Dim rCol As Range

        For i = 1 To Table.Rows.Count
            If Table.Cells(i, SearchColumnNum) = SearchValue Then
                iCount = iCount + 1
            End If

            If iCount = n Then
                VLOOKUP2 = Table.Cells(i, ResultColumnNum)
                Exit For
            End If
        Next i
End Function

'Проверка текста по шаблону
Function MaskCompare(txt As String, Mask As String, CaseSensitive As Boolean)
    If Not CaseSensitive Then
        txt = UCase(txt)
        Mask = UCase(Mask)
    End If
        
    If txt Like Mask Then
            MaskCompare = True
        Else
            MaskCompare = False
    End If
End Function

'подсчитывает количество ячеек в диапазоне, удовлетворяющих маске
Function CountByMask(rng As Range, Mask As String, CaseSensitive As Boolean)

    For Each c In rng
        If Not CaseSensitive Then
            txt = UCase(c)
            Mask = UCase(Mask)
        Else
            txt = с
        End If
        If txt Like Mask Then n = n + 1
    Next c
    CountByMask = n
End Function


'Проверка наличия в тексте символов латиницы
Function IsLatin(txt As String)
    txt = UCase(txt)
    Mask = "*[ABCDEFGHIJKLMNOPQRSTUVWXYZ]*"
        
    If txt Like Mask Then
            IsLatin = True
        Else
            IsLatin = False
    End If
End Function

'Сумма ячеек с определенным цветом заливки
Function SumByCellColor(SearchRange As Range, TargetCell As Range)
Application.Volatile True

Sum = 0

For Each cell In SearchRange
    If cell.Interior.ColorIndex = TargetCell.Interior.ColorIndex Then
        Sum = Sum + cell.value
    End If
Next
SumByCellColor = Sum
End Function

'Сумма ячеек с определенным цветом шрифта
Function SumByFontColor(SearchRange As Range, TargetCell As Range)
Application.Volatile True

Sum = 0

For Each cell In SearchRange
    If cell.Font.ColorIndex = TargetCell.Font.ColorIndex Then
        Sum = Sum + cell.value
    End If
Next
SumByFontColor = Sum
End Function

'Построение микрографиков
Function MicroCharts(rng As Range)
    Dim ChrtCodes() As Integer
    Dim outstr As String

    ReDim ChrtCodes(rng.Count)
    minval = Application.min(rng)
    minpos = Application.match(minval, rng, 0)
    maxval = Application.max(rng)
    maxpos = Application.match(maxval, rng, 0)

    If minval = 0 And maxval = 0 Then   'все нулевые значения
        For Each c In rng
            ChrtCodes(i) = 33
            i = i + 1
        Next c
        GoTo theend
    End If
    If minval >= 0 Then  'только положительные числа
        i = 0
        For Each c In rng
            ChrtCodes(i) = 68 + Round(c.value / maxval * 21)
            i = i + 1
        Next c
        GoTo theend
    End If

    If maxval <= 0 Then    ' только отрицательные числа
        i = 0
        For Each c In rng
            ChrtCodes(i) = 90 + Round(c.value / minval * 20)
            i = i + 1
        Next c
        GoTo theend
    End If

    If maxval > 0 And minval < 0 Then    'положительные и отрицательные вместе
        i = 0
        For Each c In rng
            If c.value > 0 Then
                ChrtCodes(i) = 33 + Round(c.value / maxval * 15)
            End If
            If c.value < 0 Then
                ChrtCodes(i) = 50 + Round(c.value / minval * 16)
            End If
            If c.value = 0 Then ChrtCodes(i) = 33
            i = i + 1
        Next c
    End If

theend:
    'формируем и выводим готовый массив символов
    For j = 0 To UBound(ChrtCodes)
        outstr = outstr & Chr(ChrtCodes(j))
    Next j

    MicroCharts = outstr
End Function

Function LastRow(SheetName As String) As Long

'Определение последней используемой строки на листе с именем SheetName
Dim Sh As Worksheet
Set Sh = Worksheets(SheetName)
LastRow = Sh.UsedRange.Rows.Count
LastRow = LastRow + Sh.UsedRange.Row - 1
End Function
Function LastColumn(SheetName As String, r As Long) As Range

'Определение последней используемой ячейки в строке r на листе с именем SheetName
Dim Sh As Worksheet
Dim EndCell As Range
Set Sh = Worksheets(SheetName)
Set EndCell = Sh.Cells(r, 256)
Set LastColumn = EndCell.End(xlToLeft)
End Function
Function SheetExist(SheetName As String) As Boolean
'Определение есть ли в активной книге лист с именем SheetName
Dim Sh As Object
On Error Resume Next
Set Sh = ActiveWorkbook.Worksheets(SheetName)
If Err = 0 Then SheetExist = True _
Else SheetExist = False
End Function
Function SheetExistBook(wb As Workbook, SheetName As String) As Boolean
'Определение есть ли в  книге "wb" лист с именем SheetName
Dim Sh As Object
On Error Resume Next
Set Sh = wb.Worksheets(SheetName)
If Err = 0 Then SheetExistBook = True _
Else SheetExistBook = False
End Function
Function SheetExistBookCreate(wb As Workbook, SheetName, cl As Boolean) As Worksheet
'Определение есть ли в  книге "wb" лист с именем SheetName,если нет то создает его
Dim Sh As Object
On Error Resume Next
Set Sh = wb.Worksheets(SheetName)
If Err <> 0 Then
Set Sh = wb.Worksheets.Add
Sh.name = SheetName
Else
If cl Then Sh.Cells.Clear
End If
Set SheetExistBookCreate = Sh
End Function



Function InversiaValue(v As Range)
InversiaValue = Val(Trim(v.value)) * (-1)
End Function
Function Delimeter_Count(r As Range, Delimeter As String)
'Количество Delimeter разделителей в тексте
k = 0
For i = 1 To Len(r.value)
s = Mid(r.value, i, 1)
If s = Delimeter Then k = k + 1
Next i
Delimeter_Count = k
End Function
Public Function SelectFiles(MultiSelect As Boolean, fname As String, f As String)
Dim fd As FileDialog

Set fd = Application.FileDialog(msoFileDialogFilePicker)
With fd
.InitialView = msoFileDialogViewList
.AllowMultiSelect = MultiSelect
.Filters.Clear
.Filters.Add fname, f
If .Show = -1 Then
Set SelectFiles = .SelectedItems
Else
Set SelectFiles = Nothing
End If
End With
End Function
Function ИмяЛистаВАпостров(Имя As String)
If Имя Like "* *" Then ИмяЛистаВАпостров = "'" & Имя & "'"
End Function
Function ЗаголовокСтолбца(wb As Workbook, Sh As String, NameZag As String) As Range
Dim ABS_WB As Workbook
Dim lst As Worksheet
Dim Nel_lst As Worksheet
Dim O_lst As Worksheet
Dim ABS_lst As Worksheet
Dim nr As Range
Dim f As Range

Set Nel_lst = wb.Worksheets(Sh)
Set B = LastColumn(Sh, 1)
Set zags = Nel_lst.Range(Nel_lst.Cells(1, 1), B)
Set ЗаголовокСтолбца = zags.Find(NameZag)

End Function
Function ЗаголовокСтолбцаСоздатьИлиВернуть(wb As Workbook, sSH_Name As String, sNameZag As String)
Dim lst As Worksheet
Dim zags As Range
Dim zag As Range
Set lst = wb.Worksheets(sSH_Name)
Set B = LastColumn(sSH_Name, 1)
Set zags = lst.Range(lst.Cells(1, 1), B)
Set zag = ЗаголовокСтолбца(wb, lst.name, sNameZag)
If zag Is Nothing Then
B.Offset(columnoffset:=1).value = sNameZag
Set ЗаголовокСтолбцаСоздатьИлиВернуть = B.Offset(columnoffset:=1)
Else
Set ЗаголовокСтолбцаСоздатьИлиВернуть = zag
End If
End Function

Sub grt()
 If CommentExist(ActiveCell) Then MsgBox CommentTEXT(ActiveCell)
End Sub
Function CommentExist(r As Range) As Boolean
'Определение есть ли в активной книге лист с именем SheetName
Dim Sh As Object
On Error Resume Next
Set Sh = r.Comment
If Not Sh Is Nothing Then CommentExist = True _
Else CommentExist = False
End Function

Function CommentTEXT(r As Range)
'Определение есть ли в активной книге лист с именем SheetName
Dim Sh As Object
On Error Resume Next
Set Sh = r.Comment
If Not Sh Is Nothing Then CommentTEXT = Sh.Text _

End Function

Function WorkbookExist(Path, WorkbookName As String) As Boolean
'Определение есть ли в  книге "wb" лист с именем SheetName
Dim wb As Object
On Error Resume Next
Set wb = Workbooks.Open(Path & WorkbookName)
If Err = 0 Then WorkbookExist = True _
Else WorkbookExist = False
End Function
Function IsBookOpen(wbFullName As String) As Boolean
'который проверяет открыта ли книга независимо от её
'месторасположения и используемого приложения Excel.
'Книга может быть открыта другим пользователем
'(если книга на сервере), в другом экземпляре Excel
'или в этом же экземпляре Excel.



    Dim iFF As Integer
    iFF = FreeFile
    On Error Resume Next
    Open wbFullName For Random Access Read Write Lock Read Write As #iFF
    Close #iFF
    IsBookOpen = Err
End Function
Function ОТСТУП(ячейка)
    
ОТСТУП = ячейка.IndentLevel

End Function
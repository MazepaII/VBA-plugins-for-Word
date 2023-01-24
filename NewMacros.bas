Attribute VB_Name = "NewMacros"
'Автор МАЗИЛЕВСКИЙ ИЛЬЯ ИГОРЕВИЧ
'Программы для работы соспецификациями в формате А3

Option Base 0

Dim табл()
Dim a1
Dim Лист As String
Dim колонтитулы_всех_листов, номер_активной_стр
Public Список_страниц(), подтереть_колон, Список_изм(), Список_текущ_изм(), ГОСТ_целеком
Public ГАЛКА_Проверка_и_замена_всех_колонт, ГАЛКА_Список_стр_в_файл, колонтитулы_листов_по_последн_изм As Boolean
Public ГАЛКА_Сравнение_колонтитулов_в_листах_с_указаными_в_таблице, ГАЛКА_Удаляем_разрывы As Boolean
Public ГАЛКА_Замена_точек_на_запятые, ГАЛКА_Вносим_в_колонтитулы_только_последние_изменения, ГАЛКА_Пересчет_всех_пунктов, ПОЛНЫЙ_поиск, ГАЛКА_Проверка_годов_ГОСТов As Boolean
Public Искомый_ГОСТ, Результат_ГОСТ, x_ГОСТ
Public СписокГОСТов As Collection
Public табл_у1, табл_х1, табл_ingTbIndex
Public ie As InternetExplorerMedium



Sub ПоШирене()
Attribute ПоШирене.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.ПоШирене"
'
' ПоШирене Макрос
'
'
On Error Resume Next 'пропускаем ошибки
    'Selection.Cells(1).Select
 Ячейка = Selection.Cells(1).FitText
If Ячейка = True Then
    With Selection.Cells(1)
        .Select
        .FitText = False
    End With
    With Selection.Font
        .Spacing = 0
        .Scaling = 100
        .Position = 0
        .Kerning = 0
        .Animation = wdAnimationNone
    End With
    Selection.HomeKey Unit:=wdLine 'перейти в начало строки
Else
    Selection.Cells(1).FitText = True
End If
End Sub


Sub Сумм_Масса()
'
' Сумм_Масса Макрос
'
'
On Error Resume Next 'Пропустить все ошибки


у1 = Selection.Rows.First.Index  'номер активной строки в таблице

'определение номера активной таблицы
Set tblSel = Selection.Tables(1)
ingStart = tblSel.Range.Start
For i = 1 To ActiveDocument.Tables.Count Step 1
 If ActiveDocument.Tables(i).Range.Start = ingStart Then
    ingTbIndex = i
    Exit For
 End If
 Next i
 


Количество = ActiveDocument.Tables(ingTbIndex).Cell(у1, 5).Range.Text    'снимаем значение с ячейки в таблице
Количество = Replace(Количество, Chr(13) & "", "")
Количество = Replace(Количество, " ", "")
Количество = Replace(Количество, vbrTab, "")
Количество = Replace(Количество, Chr(9), "")
Количество = Replace(Количество, ".", ",")

'Количество = Left(Количество, Len(Количество) - 1)                       'удаляем последний знак - обычно это пробел
МассаЕд = ActiveDocument.Tables(ingTbIndex).Cell(у1, 6).Range            'снимаем значение с ячейки в таблице
МассаЕд = Replace(МассаЕд, Chr(13) & "", "")
МассаЕд = Replace(МассаЕд, " ", "")
МассаЕд = Replace(МассаЕд, vbrTab, "")
МассаЕд = Replace(МассаЕд, Chr(9), "")
'МассаЕд = Left(МассаЕд, Len(МассаЕд) - 1)                                'удаляем последний знак - обычно это пробел

'If InStr(1, CStr(Количество), ".") <> 0 Or InStr(1, CStr(Количество), ",") <> 0 Then
'   количество_знаков_после_зап = Количество
'End If

'Определяем количество знаков после запятой
количество_знаков_после_зап = 0
If InStr(1, CStr(Количество), ".") <> 0 Or InStr(1, CStr(Количество), ",") <> 0 Then
   If InStr(1, CStr(Количество), ".") <> 0 Then
      количество_знаков_после_зап = Len(Split(Количество, ".")(1))
   End If
   If InStr(1, CStr(Количество), ",") <> 0 Then
      количество_знаков_после_зап = Len(Split(Количество, ",")(1))
   End If
End If

If InStr(1, CStr(МассаЕд), ".") <> 0 Or InStr(1, CStr(МассаЕд), ",") <> 0 Then
   If InStr(1, CStr(МассаЕд), ".") <> 0 Then
      If количество_знаков_после_зап < Len(Split(МассаЕд, ".")(1)) Then
         количество_знаков_после_зап = Len(Split(МассаЕд, ".")(1))
      End If
   End If
   If InStr(1, CStr(МассаЕд), ",") <> 0 Then
      If количество_знаков_после_зап < Len(Split(МассаЕд, ",")(1)) Then
         количество_знаков_после_зап = Len(Split(МассаЕд, ",")(1))
      End If
   End If
End If

'Количество = Replace(Количество, ".", Application.International(xlDecimalSeparator)) ' замена точки на запетую
'МассаЕд = Replace(МассаЕд, ".", Application.International(xlDecimalSeparator))    ' замена точки на запетую

Количество = Replace(Количество, ",", ".")  ' замена точки на запетую
МассаЕд = Replace(МассаЕд, ",", ".")    ' замена точки на запетую

If Количество = "" Or Количество = "0" Then Exit Sub
If МассаЕд = "" Or МассаЕд = "0" Then Exit Sub


Результат = Количество * МассаЕд

If Результат = 0 Then
  Количество = Replace(Количество, ".", ",")  ' замена точки на запетую
  МассаЕд = Replace(МассаЕд, ".", ",")    ' замена точки на запетую
  Результат = Количество * МассаЕд
End If
If Результат = 0 Then Результат = ""


'Округляем
'If Результат >= 0 And Результат < 0.00001 Then Результат = Format(Результат, "0.0000000")
'If Результат >= 0.00001 And Результат < 0.0001 Then Результат = Format(Результат, "0.000000")
'If Результат >= 0.0001 And Результат < 0.001 Then Результат = Format(Результат, "0.00000")
'If Результат >= 0.001 And Результат < 0.01 Then Результат = Format(Результат, "0.00000")
'If Результат >= 0.01 And Результат < 1 Then Результат = Format(Результат, "0.0000")
'If Результат >= 1 And Результат < 10 Then Результат = Format(Результат, "0.00")
'If Результат >= 10 And Результат < 500 Then Результат = Format(Результат, "0.0")
'If Результат >= 500 Then Результат = Format(Результат, "0.0")
'Результат = Round(Результат, количество_знаков_после_зап)

'Количество знаков после запятой
If количество_знаков_после_зап = 0 Then Результат = Format(Результат, "0")
If количество_знаков_после_зап = 1 Then Результат = Format(Результат, "0.0")
If количество_знаков_после_зап = 2 Then Результат = Format(Результат, "0.00")
If количество_знаков_после_зап = 3 Then Результат = Format(Результат, "0.000")
If количество_знаков_после_зап = 4 Then Результат = Format(Результат, "0.0000")
If количество_знаков_после_зап = 5 Then Результат = Format(Результат, "0.00000")
If количество_знаков_после_зап = 6 Then Результат = Format(Результат, "0.000000")
If количество_знаков_после_зап = 7 Then Результат = Format(Результат, "0.0000000")

' Заносим в ячейку
Результат = Replace(Результат, ".", ",")  ' замена запетую  на точки
ActiveDocument.Tables(ingTbIndex).Cell(у1, 7).Range = Результат

'Определяем количество знаков после запятой "Результат"
количество_знаков_после_зап_Результат = 0
If InStr(1, CStr(Результат), ".") <> 0 Or InStr(1, CStr(Результат), ",") <> 0 Then
   If InStr(1, CStr(Результат), ".") <> 0 Then
      количество_знаков_после_зап_Результат = Len(Split(Результат, ".")(1))
   End If
   If InStr(1, CStr(Результат), ",") <> 0 Then
      количество_знаков_после_зап_Результат = Len(Split(Результат, ",")(1))
   End If

'Определяем количество знаков перед запятой "Результат"
количество_знаков_перед_зап_Результат = 0
   If InStr(1, CStr(Результат), ".") <> 0 Then
      количество_знаков_перед_зап_Результат = Len(Split(Результат, ".")(0))
   End If
   If InStr(1, CStr(Результат), ",") <> 0 Then
      количество_знаков_перед_зап_Результат = Len(Split(Результат, ",")(0))
   End If
Else
    количество_знаков_перед_зап_Результат = Len(Результат)
End If


If количество_знаков_после_зап_Результат + количество_знаков_перед_зап_Результат > 7 Then
ActiveDocument.Tables(ingTbIndex).Cell(у1, 7).Range.Select
Call ПоШирене
End If

End Sub
Sub Фиолетовый()
Attribute Фиолетовый.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Фиолетовый"
'
' Фиолетовый Макрос
'
   If Selection.Font.Color <> 10498160 Then
    Selection.Font.Color = 10498160
    Selection.Font.Bold = True
    Exit Sub
   End If
    
    If Selection.Font.Color = 10498160 Then
      Selection.Font.Color = wbblack
      Selection.Font.Bold = False
      Exit Sub
    End If
End Sub



Function ArrayOfValues(ByVal Txt$) As Variant
 
' Принимает в качестве параметра строку типа ",,5,6,8,,9-15,18,2,11-9,,1,4,,21,"
    ' Возвращает одномерный (горизонтальный) массив в формате
    ' array(5,6,8,9,10,11,12,13,14,15,18,2,11,10,9,1,4,21)
    ' (пустые значения удаляются; диапазоны типа 9-15 и 17-13 раскрываются)

arr = Split(Replace(Txt$, " ", ""), ","): Dim n As Long: ReDim tmpArr(0 To 0)
    For i = LBound(arr) To UBound(arr)
        Select Case True
            Case arr(i) = "", Val(arr(i)) < 0
                '  раскомментируйте эту строку, чтобы пустые и нулевые значения
                '  тоже добавлялись в результат (преобразовывались в значение -1)
                'tmpArr(UBound(tmpArr)) = -1: ReDim Preserve tmpArr(0 To UBound(tmpArr) + 1)
            Case IsNumeric(arr(i))
                tmpArr(UBound(tmpArr)) = arr(i): ReDim Preserve tmpArr(0 To UBound(tmpArr) + 1)
            Case arr(i) Like "*#-#*"
                spl = Split(arr(i), "-")
                If UBound(spl) = 1 Then
                    If IsNumeric(spl(0)) And IsNumeric(spl(1)) Then
                        For j = Val(spl(0)) To Val(spl(1)) Step IIf(Val(spl(0)) > Val(spl(1)), -1, 1)
                            tmpArr(UBound(tmpArr)) = j: ReDim Preserve tmpArr(0 To UBound(tmpArr) + 1)
                        Next j
                    End If
                End If
        End Select
    Next i
    On Error Resume Next: ReDim Preserve tmpArr(0 To UBound(tmpArr) - 1)
    ArrayOfValues = tmpArr
End Function
 

Sub Заносим_в_колонтитул()
t = Timer
    
If колонтитулы_листов_по_последн_изм = True Then
 GoTo колонтитулы_листов_по_последн
End If
    
  If колонтитулы_всех_листов = False Then      'не работает макрос Берем_намера_стр_из_табл
     сообщение = MsgBox("Yes-""Зам."", No-""Нов."", Cancel-выйти ", vbYesNoCancel + vbQuestion + vbDefaultButton1, "Запрос на список страниц")
     If сообщение = vbCancel Then Exit Sub
     If сообщение = vbYes Then Лист = "Зам."
     If сообщение = vbNo Then Лист = "Нов."
  End If
  
колонтитулы_листов_по_последн:
    
 If колонтитулы_всех_листов = False Then
       Application.ScreenUpdating = False 'отключить обновление документа
 End If

 If Selection.PageSetup.PageWidth = CSng(Format(CentimetersToPoints(21), "0.0")) And _
      Selection.PageSetup.PageHeight = CSng(Format(CentimetersToPoints(29.7), "0.0")) And _
      Selection.PageSetup.VerticalAlignment = 0 _
 Then
      Формат_листа = "А4"
 Else
      Формат_листа = "А3"
 End If

    
If Формат_листа = "А4" Then
   текущий_масштаб = ActiveWindow.ActivePane.View.Zoom.Percentage
   ActiveWindow.ActivePane.View.Zoom.Percentage = 150
End If
    

    'Application.ScreenUpdating = False 'отключить обновление документа
' разбиваем строку в массив, содержащий все значения исходной строки
    
'+++++Получем список страниц докума
'путь = ActiveDocument.Path
'имя = ActiveDocument.Name
'Open путь & "\" & имя & "_страницы.txt" For Input As #1  'открыть для чтения "Input"
'Line Input #1, текст_из_txt 'чтение строки #1
'Close #1 'закрыть документ

'Массив = Split(текст_из_txt, ",") 'получаем массив
'ReDim Список_страниц(UBound(Массив, 1)) 'задаем размер таблицы
'присваеваем значения оного массива другому

'For i = 0 To UBound(Массив, 1)
'   Список_страниц(i) = Массив(i)
'Next i
   'номер активной стр не в колонтитуле
If колонтитулы_листов_по_последн_изм = False Then 'чтобы не заглючело и невыходить из колонтитула чтобы получить стр
   номер_активной_стр = Selection.Information(wdActiveEndPageNumber) 'номер активной стр
End If

      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
      WordBasic.ViewFooterOnly 'перейти на нижний ка лантитул
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
      WordBasic.ViewFooterOnly 'перейти на нижний калантитул
'If колонтитулы_листов_по_последн_изм = False Then

'End If
      ' номер_листа_взади_стар
      ' номер_листа_искомый_стар
      ' номер_листа_впереди_стар
' If колонтитулы_всех_листов = False Then     'не работает макрос Берем_намера_стр_из_табл
      Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=номер_активной_стр - 1, Name:="" 'перейти на стр по номеру
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
      WordBasic.ViewFooterOnly 'перейти на нижний ка лантитул
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
      WordBasic.ViewFooterOnly 'перейти на нижний калантитул
      If Selection.Information(wdActiveEndPageNumber) = 1 Then
         номер_листа_взади_стар = "1"
      Else
         номер_листа_взади_стар = Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text
         номер_листа_взади_стар = Replace(номер_листа_взади_стар, Chr(13) & "", "")
      End If
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
      
      Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=номер_активной_стр, Name:="" 'перейти на стр по номеру
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
      WordBasic.ViewFooterOnly 'перейти на нижний калантитул
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
      WordBasic.ViewFooterOnly 'перейти на нижний калантитул
      номер_листа_искомый_стар = Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text
      номер_листа_искомый_стар = Replace(номер_листа_искомый_стар, Chr(13) & "", "")
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
      
      Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=номер_активной_стр + 1, Name:="" 'перейти на стр по номеру
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
      WordBasic.ViewFooterOnly 'перейти на нижний калантитул
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
      WordBasic.ViewFooterOnly 'перейти на нижний калантитул
      номер_листа_впереди_стар = Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text
      номер_листа_впереди_стар = Replace(номер_листа_впереди_стар, Chr(13) & "", "")
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
'End If

' If колонтитулы_всех_листов = True Then     'работает макрос Берем_намера_стр_из_табл
'       номер_листа_взади_стар = Список_страниц(номер_активной_стр - 2)
'       номер_листа_искомый_стар = Список_страниц(номер_активной_стр - 1)
'       номер_листа_впереди_стар = Список_страниц(номер_активной_стр)
' End If

'если False заполняем толко текущую стр
 If колонтитулы_всех_листов = False Then     'не работает макрос Берем_намера_стр_из_табл
   Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=номер_активной_стр, Name:="" 'перейти на стр по номеру
   With ActiveDocument.Range.Tables(ActiveDocument.Range.Tables.Count)
      Erase табл 'отчистить массив
      ReDim табл(.Rows.Count * 2, .Rows.Count * 2) 'задаем размер таблицы
      
      If Формат_листа = "А4" Then
         номер_строки_старт = 3
      Else
         номер_строки_старт = 4
      End If
      'последнее изменение
      For i = номер_строки_старт To .Rows.Count  ' получение стр из последней табл для каждого изм (а1, а2 ...)
         If .Cell(i, 1).Range.Text <> Chr(13) & "" Then
            табл(a1, 1) = Left(.Cell(i, 1).Range.Text, Len(.Cell(i, 1).Range.Text) - 2)
            'Лист = "Зам."
            табл(a1, 4) = Left(.Cell(i, 7).Range.Text, Len(.Cell(i, 7).Range.Text) - 2)
            табл(a1, 5) = Left(.Cell(i, 9).Range.Text, Len(.Cell(i, 9).Range.Text) - 2)
            табл(a1, 6) = Left(.Cell(i, 10).Range.Text, Len(.Cell(i, 10).Range.Text) - 2)
         End If
      Next i
   End With
 End If
 

'+++Вставляем разрыв страницы если его нет и снимаем галку "как в предыдущем разделе -  страница за текущей"

    
    Selection.MoveUp Unit:=wdLine, Count:=1, Extend:=wdExtend

    
    Selection.Find.ClearFormatting
    f = False
    With Selection.Find
        .Text = "^b"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop 'wdFindStop не спарашивать если не нашло
         f = Selection.Find.Execute
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    If f = False Then
      Selection.MoveDown Unit:=wdLine, Count:=1
      If Формат_листа = "А4" Then
         Selection.HomeKey Unit:=wdLine 'перейти в начало строки
         Selection.InsertBreak Type:=wdSectionBreakNextPage    'разрыв на следующей странице
      Else
         Selection.MoveUp Unit:=wdLine, Count:=1
         Selection.InsertBreak Type:=wdSectionBreakContinuous  'разрыв на текущей странице на следующей стр
      End If
    Else
      Selection.MoveRight Unit:=wdCharacter, Count:=1
    End If
    
'+++Вставляем разрыв страницы если его нет и снимаем галку "как в предыдущем разделе -  страница перед текущей"
      'снимаем номер следуйщей стр
    Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=номер_активной_стр + 1, Name:="" 'перейти на стр по номеру
    Selection.MoveUp Unit:=wdLine, Count:=1, Extend:=wdExtend
    Selection.Find.ClearFormatting
    f = False
    With Selection.Find
        .Text = "^b"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop 'wdFindStop не спарашивать если не нашло
         f = Selection.Find.Execute
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    If f = False Then
      Selection.MoveDown Unit:=wdLine, Count:=1
      If Формат_листа = "А4" Then
         Selection.HomeKey Unit:=wdLine 'перейти в начало строки
         Selection.InsertBreak Type:=wdSectionBreakNextPage    'разрыв на следующей странице
      Else
         Selection.MoveUp Unit:=wdLine, Count:=1
         Selection.InsertBreak Type:=wdSectionBreakContinuous  'разрыв на текущей странице на следующей стр
      End If
    Else
     Selection.MoveRight Unit:=wdCharacter, Count:=1
    End If
      Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=номер_активной_стр + 1, Name:="" 'перейти на стр по номеру
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
      WordBasic.ViewFooterOnly 'перейти на нижний калантитул
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
      WordBasic.ViewFooterOnly 'перейти на нижний калантитул
      'WordBasic.ViewFooterOnly 'перейти на нижний калантитул
      Selection.HeaderFooter.LinkToPrevious = False ' снять "как в педыдущем разделе"
      номер_листа_впереди_нов = Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text
      номер_листа_впереди_нов = Replace(номер_листа_впереди_нов, "", "")
      номер_листа_впереди_нов = Replace(номер_листа_впереди_нов, Chr(13), "")
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула

      Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=номер_активной_стр, Name:="" 'перейти на стр по номеру
      WordBasic.ViewFooterOnly ' открыть нижний колонтитул
      
      номер_листа_искомый_нов = Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text
      номер_листа_искомый_нов = Replace(номер_листа_искомый_нов, "", "")
      номер_листа_искомый_нов = Replace(номер_листа_искомый_нов, Chr(13), "")
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
      
     'снимаем номер предыдущей стр
    If номер_активной_стр <> номер_активной_стр Then
      Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=номер_активной_стр - 1, Name:="" 'перейти на стр по номеру
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
      WordBasic.ViewFooterOnly 'перейти на нижний калантитул
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
      WordBasic.ViewFooterOnly 'перейти на нижний калантитул
      номер_листа_взади_нов = Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text
      номер_листа_взади_нов = Replace(номер_листа_взади_нов, "", "")
      номер_листа_взади_нов = Replace(номер_листа_взади_нов, Chr(13), "")
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
      Selection.GoTo What:=wdGoToPage, Which:=wdGoToNext, Name:="+1"
    End If
      
      ' номер_листа_взади_нов
      ' номер_листа_искомый_нов
      ' номер_листа_впереди_нов
      
      ' номер_листа_взади_стар
      ' номер_листа_искомый_стар
      ' номер_листа_впереди_стар

      If номер_листа_взади_нов <> номер_листа_взади_стар And номер_активной_стр <> номер_активной_стр Then
            Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=номер_активной_стр - 1, Name:="" 'перейти на стр по номеру
           WordBasic.ViewFooterOnly ' открыть нижний колонтитул
             With Selection.HeaderFooter.PageNumbers  ' продолжить нумацию
               .NumberStyle = wdPageNumberStyleArabic
               .HeadingLevelForChapter = 0
               .IncludeChapterNumber = False
               .ChapterPageSeparator = wdSeparatorHyphen
               .RestartNumberingAtSection = False
               .StartingNumber = 0
            End With
            If Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text, Chr(13) & "", "") <> номер_листа_взади_стар Then
               With Selection.HeaderFooter.PageNumbers  ' задать номер стр номером
                 .NumberStyle = wdPageNumberStyleArabic
                 .HeadingLevelForChapter = 0
                 .IncludeChapterNumber = False
                 .ChapterPageSeparator = wdSeparatorHyphen
                 .RestartNumberingAtSection = True
                 .StartingNumber = номер_листа_взади_стар
              End With
           End If
           ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
      End If
      
      If номер_листа_искомый_нов <> номер_листа_искомый_стар Then
            Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=номер_активной_стр, Name:="" 'перейти на стр по номеру
           WordBasic.ViewFooterOnly ' открыть нижний колонтитул
             With Selection.HeaderFooter.PageNumbers  ' продолжить нумацию
               .NumberStyle = wdPageNumberStyleArabic
               .HeadingLevelForChapter = 0
               .IncludeChapterNumber = False
               .ChapterPageSeparator = wdSeparatorHyphen
               .RestartNumberingAtSection = False
               .StartingNumber = 0
            End With
            If Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text, Chr(13) & "", "") <> номер_листа_искомый_стар Then
               With Selection.HeaderFooter.PageNumbers  ' задать номер стр номером
                 .NumberStyle = wdPageNumberStyleArabic
                 .HeadingLevelForChapter = 0
                 .IncludeChapterNumber = False
                 .ChapterPageSeparator = wdSeparatorHyphen
                 .RestartNumberingAtSection = True
                 .StartingNumber = Список_страниц(номер_активной_стр - 1)
              End With
           End If
           If Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text, Chr(13) & "", "") <> номер_листа_искомый_стар Then
               Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text = номер_листа_искомый_стар
           End If
           ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
      End If
      
      
      If номер_листа_впереди_нов <> номер_листа_впереди_стар Then
            Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=номер_активной_стр + 1, Name:="" 'перейти на стр по номеру
           WordBasic.ViewFooterOnly ' открыть нижний колонтитул
             With Selection.HeaderFooter.PageNumbers  ' продолжить нумацию
               .NumberStyle = wdPageNumberStyleArabic
               .HeadingLevelForChapter = 0
               .IncludeChapterNumber = False
               .ChapterPageSeparator = wdSeparatorHyphen
               .RestartNumberingAtSection = False
               .StartingNumber = 0
            End With
            If Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text, Chr(13) & "", "") <> номер_листа_впереди_стар Then
               With Selection.HeaderFooter.PageNumbers  ' задать номер стр номером
                 .NumberStyle = wdPageNumberStyleArabic
                 .HeadingLevelForChapter = 0
                 .IncludeChapterNumber = False
                 .ChapterPageSeparator = wdSeparatorHyphen
                 .RestartNumberingAtSection = True
                 .StartingNumber = номер_листа_впереди_стар
              End With
           End If
           ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
      End If

      
'++++++++Текущая страница

    'ActiveWindow.ActivePane.View.NextHeaderFooter 'следующий раздел
    'ActiveWindow.ActivePane.View.PreviousHeaderFooter 'предыдущий раздел
     Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=номер_активной_стр, Name:="" 'перейти на стр по номеру
     
   WordBasic.ViewFooterOnly ' открыть нижний колонтитул
    If Selection.HeaderFooter.LinkToPrevious = True Then ' снять "как в педыдущем разделе"
       Selection.HeaderFooter.LinkToPrevious = False ' снять "как в педыдущем разделе"
       ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула если не выйти из колонтитула и сново не войти глючит и вписывает текст не втот колонтитул
       WordBasic.ViewFooterOnly ' открыть нижний колонтитул
    End If

  
'ReDim табл(1, 6) 'задаем размер таблицы
'табл(a1, 1) = "а1"
'Лист = "Зам."
'табл(a1, 4) = "22220.43.___"
'табл(a1, 5) = "Мазилевский"
'табл(a1, 6) = "15.16.18"

'заполняем колонтитул текстом
  If подтереть_колон = False Then
   '   Selection.Find.Execute 'выделить начденый текст  'Dim табл()
      'Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 1).Range.Text = табл(a1, 1) 'таблица в колонтитуле
      'Selection.Tables(1).Cell(2, 1).Range.Text = "а1" 'таблица в колонтитуле
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 2).Range.Text = Лист 'таблица в колонтитуле
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 3).Range.Text = табл(a1, 4) 'таблица в колонтитуле
      'Selection.Tables(1).Cell(2, 3).Range.Text = "22220.43.___" 'таблица в колонтитуле
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 3).Select
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 3).FitText = True
      'Selection.HeaderFooter.Range.Cells(1).FitText = True  'сузить до ширены ячейки
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 4).Range.Text = табл(a1, 5) 'таблица в колонтитуле
      'Selection.Tables(1).Cell(2, 4).Range.Text = "Мазилевский" 'таблица в колонтитуле
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 4).Select
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 4).FitText = True  'сузить до ширены ячейки
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 5).Range.Text = табл(a1, 6) 'таблица в колонтитуле
      'Selection.Tables(1).Cell(2, 5).Range.Text = "15.16.18" 'таблица в колонтитуле
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 5).Select
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 5).FitText = True
  End If
  
'делаем колонтитул пустым
  If подтереть_колон = True Then
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 1).Range.Text = "" 'таблица в колонтитуле
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 2).Range.Text = "" 'таблица в колонтитуле
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 3).Range.Text = "" 'таблица в колонтитуле
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 4).Range.Text = "" 'таблица в колонтитуле
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 5).Range.Text = "" 'таблица в колонтитуле
  End If

 
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
  
    
    
    If (Timer - t) > 60 Then
      Debug.Print (Timer - t) / 60 & " мин" ' время  в сек
      Else
      Debug.Print Timer - t & " сек" ' время  в сек
    End If
    
    If колонтитулы_всех_листов = True Or колонтитулы_листов_по_последн_изм = True Then
       GoTo к_выходу
    End If
    
If Формат_листа = "А4" Then
    ActiveWindow.ActivePane.View.Zoom.Percentage = текущий_масштаб
End If

    
    Application.ScreenUpdating = True 'включить обновление документа
    
к_выходу:
    
    'Application.ScreenUpdating = True 'включить обновление документа
End Sub


Sub Список_стр_в_файл()
'Dim Список_страниц()
  t = Timer
  
  Application.ScreenUpdating = False 'отключить обновление документа
  
путь = ActiveDocument.Path
имя = ActiveDocument.Name
Имя_файла = имя & "_страницы.txt"



'сообщение = MsgBox("Обновить/создать список страниц в файле:" & Chr(13) & """" & Имя_файла & """", vbYesNoCancel + vbQuestion + vbDefaultButton1, "Запрос на список страниц")

If сообщение = vbCancel Then
  Application.ScreenUpdating = True
  Exit Sub
End If


  
  'Номер таблицы с деталюхами
For i = 1 To ActiveDocument.Tables.Count Step 1
   If ActiveDocument.Tables(i).Range.Columns.Count = 18 Then
      Номер_таблицы = i
      Exit For
   End If
Next i

ActiveDocument.Tables(Номер_таблицы).Range.Cells(1).Select
номер_стр_с_первой_табл = Selection.Information(wdActiveEndPageNumber) 'номер активной стр
 
'If сообщение = vbYes Then
  Erase Список_страниц 'отчистить массив
  ReDim Список_страниц(номер_стр_с_первой_табл) 'задаем размер таблицы
  
  For i = 0 To номер_стр_с_первой_табл - 2
     Список_страниц(i) = Str(i + 1)
  Next i
  
  Колич_листов_в_докум = ActiveDocument.ComputeStatistics(wdStatisticPages)
  Selection.HomeKey Unit:=wdStory 'переход в начало документа
  For j = номер_стр_с_первой_табл - 1 To ActiveDocument.ComputeStatistics(wdStatisticPages) - 1 'количество страниц
       
       номер_активной_стр = Selection.Information(wdActiveEndPageNumber) 'номер активной стр
       DoEvents
       UserForm1.Label1.Width = CInt((300 * номер_активной_стр) / Колич_листов_в_докум)
       UserForm1.Label_Лист = "Лист " & номер_активной_стр & " из " & Колич_листов_в_докум
       'If (Timer - t) > 60 Then
       '   UserForm1.Label2 = "Время " & Round((Timer - t) / 60, 1) & " мин" ' время  в сек
       'Else
       '   UserForm1.Label2 = "Время " & Round(Timer - t, 1) & " сек" ' время  в сек
       'End If
       UserForm1.Label2 = "Время: " & TimeSerial(0, 0, Timer - t)
       
       UserForm1.Repaint
              
      Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=j + 1, Name:="" 'перейти на стр по номеру
      WordBasic.ViewFooterOnly ' открыть нижний колонтитул
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
      WordBasic.ViewFooterOnly ' открыть нижний колонтитул
      ReDim Preserve Список_страниц(j)
      Список_страниц(j) = Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text 'обращение к ячейке в табл в нижнем колонтитуле
      Список_страниц(j) = Replace(Список_страниц(j), "", "")
      Список_страниц(j) = Replace(Список_страниц(j), Chr(13), "")
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
  Next j
'End If 'If сообщение = vbYes Then


'проверка существует ли файл
'If Dir(путь & "\" & Имя_файла) = "" Then
'создает файл (если файл уже есть то переписывает на пустой)
'    Set fso = CreateObject("scripting.filesystemobject")
'    Set ts = fso.createtextfile(путь & "\" & имя & "_страницы.txt", True)
'    ts.write txt: ts.Close
'    Set ts = Nothing: Set fso = Nothing
'End If

'открыть и создает имееющейся файл и сохранить
    Open путь & "\" & Имя_файла For Output As #1
    Print #1, Join(Список_страниц, ",")
    Close #1


    Application.ScreenUpdating = True 'включить обновление документа
    'End If
    If (Timer - t) > 60 Then
      Debug.Print (Timer - t) / 60 & " мин" ' время  в сек
      Else
      Debug.Print Timer - t & " сек" ' время  в сек
    End If
End Sub




Sub Удалить_ленту()

    '''''востановить сочетания клавиш по умолчанию
    CustomizationContext = NormalTemplate
    KeyBindings.ClearAll

'Удаляю все кнопки
On Error Resume Next 'пропускаем ошибки
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
On Error GoTo 0  'сново не пропускаем ошибки
End Sub





Sub Получить_список_изм_постранично()
Dim Массив_стр_изм(), Массив_стр_зам(), Массив_стр_нов
Dim стр_изм, стр_зам, стр_нов 'текст



  
  номер_активной_стр = Selection.Information(wdActiveEndPageNumber) 'номер активной стр
  If номер_активной_стр = 1 Then Exit Sub
  
  'снимаем  номер активной страницы
  WordBasic.ViewFooterOnly ' открыть нижний колонтитул
  номер_активной_стр_из_колонт = Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text, Chr(13) & "", "")
  ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
  
  Колич_листов_в_докум = ActiveDocument.ComputeStatistics(wdStatisticPages)
  
  With ActiveDocument.Range.Tables(ActiveDocument.Range.Tables.Count)
    Erase Список_изм 'отчистить массив
    'ReDim Список_изм(.Rows.Count * 2, Колич_листов_в_докум, Колич_листов_в_докум, Колич_листов_в_докум, 1, 1, 1) 'задаем размер таблицы
    ReDim Список_изм(Колич_листов_в_докум * 3, 5) 'задаем размер таблицы  листов в 3 раза больше вдруг ктото анулирует лист это происходи крайне редко
    'Список_страниц(0) = "1"
 'список изменений
 колич_изм = 0
 For i = 4 To .Rows.Count  ' получение стр из последней табл для каждого изм (а1, а2 ...)
    '.Cell(i, 1).Select
    If .Cell(i, 1).Range.Text <> Chr(13) & "" Then
        колич_изм = колич_изм + 1
    End If
 Next i
 
    номер_измен = 0
    стр_изм = ""
    стр_зам = ""
    стр_нов = ""
    g = 1
      For i = 4 To .Rows.Count  ' получение стр из последней табл для каждого изм (а1, а2 ...)
         '.Cell(i, 1).Select
        
         If .Cell(i, 1).Range.Text <> Chr(13) & "" Or .Cell(i, 2).Range.Text <> Chr(13) & "" _
         Or .Cell(i, 3).Range.Text <> Chr(13) & "" Or .Cell(i, 4).Range.Text <> Chr(13) & "" Then
           '.Cell(i, 1).Select
           
           If .Cell(i, 1).Range.Text <> Chr(13) & "" Then
               номер_измен = номер_измен + 1
           End If
           
           
           If .Cell(i, 1).Range.Text <> Chr(13) & "" Then
               текущ_изм = Left(.Cell(i, 1).Range.Text, Len(.Cell(i, 1).Range.Text) - 2)
           End If
               стр_изм = стр_изм + Left(.Cell(i, 2).Range.Text, Len(.Cell(i, 2).Range.Text) - 2)
               стр_зам = стр_зам + Left(.Cell(i, 3).Range.Text, Len(.Cell(i, 3).Range.Text) - 2)
               стр_нов = стр_нов + Left(.Cell(i, 4).Range.Text, Len(.Cell(i, 4).Range.Text) - 2)
               Номер_докум = Номер_докум + Left(.Cell(i, 7).Range.Text, Len(.Cell(i, 7).Range.Text) - 2)
               Подпись = Подпись + Left(.Cell(i, 9).Range.Text, Len(.Cell(i, 9).Range.Text) - 2)
               Дата = Дата + Left(.Cell(i, 10).Range.Text, Len(.Cell(i, 10).Range.Text) - 2)
         End If
         
         
         If (.Cell(i + 1, 1).Range.Text <> Chr(13) & "" And номер_измен <> 0) Or (i = .Rows.Count) Then  'без Or (i = .Rows.Count) заносит все кроме последнего изм
               'переводим строчный список стр в масивый (прикаждом оращении предыдущий список обнуляется)
               Массив_стр_изм = ArrayOfValues(стр_изм)
               Массив_стр_зам = ArrayOfValues(стр_зам)
               Массив_стр_нов = ArrayOfValues(стр_нов)
               'делаем массивый одной длины
               If UBound(Массив_стр_изм, 1) >= UBound(Массив_стр_зам, 1) And UBound(Массив_стр_изм, 1) >= UBound(Массив_стр_нов, 1) Then
                 ReDim Preserve Массив_стр_зам(UBound(Массив_стр_изм, 1))
                 ReDim Preserve Массив_стр_нов(UBound(Массив_стр_изм, 1))
               End If
               If UBound(Массив_стр_зам, 1) >= UBound(Массив_стр_изм, 1) And UBound(Массив_стр_зам, 1) >= UBound(Массив_стр_нов, 1) Then
                 ReDim Preserve Массив_стр_изм(UBound(Массив_стр_зам, 1))
                 ReDim Preserve Массив_стр_нов(UBound(Массив_стр_зам, 1))
               End If
               If UBound(Массив_стр_нов, 1) >= UBound(Массив_стр_изм, 1) And UBound(Массив_стр_нов, 1) >= UBound(Массив_стр_зам, 1) Then
                 ReDim Preserve Массив_стр_изм(UBound(Массив_стр_нов, 1))
                 ReDim Preserve Массив_стр_зам(UBound(Массив_стр_нов, 1))
               End If
               
               For Y = 0 To (UBound(Массив_стр_изм, 1))
                  On Error Resume Next
                  Массив_стр_изм(Y) = CSng(Массив_стр_изм(Y)) 'переводим  число с запятой (удаляем пробелы)
                  Массив_стр_изм(Y) = CStr(Массив_стр_изм(Y)) 'переводим  в текст
                  Массив_стр_изм(Y) = Replace(Массив_стр_изм(Y), ",", ".")
                  Массив_стр_зам(Y) = CSng(Массив_стр_зам(Y)) 'переводим  число с запятой (удаляем пробелы)
                  Массив_стр_зам(Y) = CStr(Массив_стр_зам(Y)) 'переводим  в текст
                  Массив_стр_зам(Y) = Replace(Массив_стр_зам(Y), ",", ".")
                  Массив_стр_нов(Y) = CSng(Массив_стр_нов(Y)) 'переводим  число с запятой (удаляем пробелы)
                  Массив_стр_нов(Y) = CStr(Массив_стр_нов(Y)) 'переводим  в текст
                  Массив_стр_нов(Y) = Replace(Массив_стр_нов(Y), ",", ".")
                  On Error GoTo 0
               Next Y
               'Заносим список изменений для каждой страницы в отдельности в массив
               'номер_стр = 0
               'Массив_стр_изм
               For Y = 0 To (UBound(Массив_стр_изм, 1))
                  If Массив_стр_изм(Y) <> 0 Then
                     нашло = False
                     'Перебераем список в поисках страницы с темже номером
                     For h = 0 To (UBound(Список_изм, 1))
                        If Список_изм(h, 0) = Массив_стр_изм(Y) Then
                           Список_изм(h, 0) = Массив_стр_изм(Y)
                           Список_изм(h, 1) = текущ_изм
                           Список_изм(h, 2) = ""
                           Список_изм(h, 3) = Номер_докум
                           Список_изм(h, 4) = Подпись
                           Список_изм(h, 5) = Дата
                           'номер_стр = номер_стр + 1
                           нашло = True 'нашло совпадение
                           Exit For 'нашло и занесло выйти из цикла прогона
                        End If
                     Next h
                        If нашло = False Then 'не нашло совпадение
                           Список_изм(номер_стр, 0) = Массив_стр_изм(Y)
                           Список_изм(номер_стр, 1) = текущ_изм
                           Список_изм(номер_стр, 2) = ""
                           Список_изм(номер_стр, 3) = Номер_докум
                           Список_изм(номер_стр, 4) = Подпись
                           Список_изм(номер_стр, 5) = Дата
                           номер_стр = номер_стр + 1
                        End If
                  End If
               Next Y
               'Массив_стр_зам
               For Y = 0 To (UBound(Массив_стр_зам, 1))
                  If Массив_стр_зам(Y) <> 0 Then
                     нашло = False
                     'Перебераем список в поисках страницы с темже номером
                     For h = 0 To (UBound(Список_изм, 1))
                        If Список_изм(h, 0) = Массив_стр_зам(Y) Then
                           Список_изм(h, 0) = Массив_стр_зам(Y)
                           Список_изм(h, 1) = текущ_изм
                           Список_изм(h, 2) = "Зам."
                           Список_изм(h, 3) = Номер_докум
                           Список_изм(h, 4) = Подпись
                           Список_изм(h, 5) = Дата
                           'номер_стр = номер_стр + 1
                           нашло = True 'нашло совпадение
                           Exit For
                        End If
                     Next h
                        If нашло = False Then 'не нашло совпадение
                           Список_изм(номер_стр, 0) = Массив_стр_зам(Y)
                           Список_изм(номер_стр, 1) = текущ_изм
                           Список_изм(номер_стр, 2) = "Зам."
                           Список_изм(номер_стр, 3) = Номер_докум
                           Список_изм(номер_стр, 4) = Подпись
                           Список_изм(номер_стр, 5) = Дата
                           номер_стр = номер_стр + 1
                        End If
                  End If
               Next Y
               'Массив_стр_нов
               For Y = 0 To (UBound(Массив_стр_нов, 1))
                  If Массив_стр_нов(Y) <> 0 Then
                     нашло = False
                     'Перебераем список в поисках страницы с темже номером
                     For h = 0 To (UBound(Список_изм, 1))
                        If Список_изм(h, 0) = Массив_стр_нов(Y) Then
                           Список_изм(h, 0) = Массив_стр_нов(Y)
                           Список_изм(h, 1) = текущ_изм
                           Список_изм(h, 2) = "Нов."
                           Список_изм(h, 3) = Номер_докум
                           Список_изм(h, 4) = Подпись
                           Список_изм(h, 5) = Дата
                           'номер_стр = номер_стр + 1
                           нашло = True 'нашло совпадение
                           Exit For
                        End If
                     Next h
                        If нашло = False Then 'не нашло совпадение
                           Список_изм(номер_стр, 0) = Массив_стр_нов(Y)
                           Список_изм(номер_стр, 1) = текущ_изм
                           Список_изм(номер_стр, 2) = "Нов."
                           Список_изм(номер_стр, 3) = Номер_докум
                           Список_изм(номер_стр, 4) = Подпись
                           Список_изм(номер_стр, 5) = Дата
                           номер_стр = номер_стр + 1
                        End If
                  End If
               Next Y
           'Обнуляем
           стр_изм = ""
           стр_зам = ""
           стр_нов = ""
           Номер_докум = ""
           Подпись = ""
           Дата = ""
           End If
         
      Next i
   End With
  

 
End Sub

Sub Проверка_и_замена_всех_колонт()
  Время = Timer
  
  DoEvents
  
'Номер таблицы с деталюхами
For i = 1 To ActiveDocument.Tables.Count Step 1
   If ActiveDocument.Tables(i).Range.Columns.Count = 18 Then
      Номер_таблицы = i
      Exit For
   End If
Next i

ActiveDocument.Tables(Номер_таблицы).Range.Cells(1).Select
номер_стр_с_первой_табл = Selection.Information(wdActiveEndPageNumber) 'номер активной стр
   
 
  'UserForm1.Label1.Width = CInt((300 * номер_активной_стр) / Колич_листов_в_докум)
  'UserForm1.Label_Лист = "Лист " & номер_активной_стр & " из " & Колич_листов_в_докум
  UserForm1.Label1.Width = 0
  UserForm1.Repaint
  
 
  
  Application.ScreenUpdating = False 'отключить обновление документа
  'UserForm1.Label1.Width = 0
  'UserForm1.Show
    Колич_листов_в_докум = ActiveDocument.ComputeStatistics(wdStatisticPages)
    With ActiveDocument.Range.Tables(ActiveDocument.Range.Tables.Count)
       Erase Список_изм 'отчистить массив
      'ReDim Список_изм(.Rows.Count * 2, Колич_листов_в_докум, Колич_листов_в_докум, Колич_листов_в_докум, 1, 1, 1) 'задаем размер таблицы
       ReDim Список_изм(Колич_листов_в_докум * 3, 5) 'задаем размер таблицы
    End With
    
Call Получить_список_изм_постранично  'Список_изм()
  
  With ActiveDocument.Range.Tables(ActiveDocument.Range.Tables.Count)
      Erase табл 'отчистить массив
      ReDim табл(.Rows.Count * 2, .Rows.Count * 2) 'задаем размер таблицы
  End With
  a1 = 0
  

  For i = 2 To Колич_листов_в_докум

       'Пропускаем листы с содержанием
       If i < номер_стр_с_первой_табл Then i = номер_стр_с_первой_табл
       
       Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=i, Name:="" 'перейти на стр по номеру
       WordBasic.ViewFooterOnly ' открыть нижний колонтитул
       ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
       WordBasic.ViewFooterOnly ' открыть нижний колонтитул
       номер_активной_стр_из_колонт = Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text, Chr(13) & "", "")
       изм_активной_стр_из_колонт = Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 1).Range.Text, Chr(13) & "", "")
       лист_активной_стр_из_колонт = Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 2).Range.Text, Chr(13) & "", "")
       номер_докум_активной_стр_из_колонт = Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 3).Range.Text, Chr(13) & "", "")
       
       ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
       
       номер_активной_стр = Selection.Information(wdActiveEndPageNumber) 'номер активной стр
       'Application.ScreenUpdating = True
       'Application.StatusBar = "Лист " & номер_активной_стр & " из " & Колич_листов_в_докум
       'Application.StatusBar = False
       'Application.ScreenUpdating = False
       'Application.ScreenUpdating = True 'включить обновление документа
       DoEvents
       UserForm1.Label1.Width = CInt((300 * номер_активной_стр) / Колич_листов_в_докум)
       UserForm1.Label_Лист = "Лист " & номер_активной_стр & " из " & Колич_листов_в_докум
       'If (Timer - Время) > 60 Then
       '   UserForm1.Label2 = "Время " & Round((Timer - Время) / 60, 1) & " мин" ' время  в сек
       'Else
       '   UserForm1.Label2 = "Время " & Round(Timer - Время, 1) & " сек" ' время  в сек
       'End If
       UserForm1.Label2 = "Время: " & TimeSerial(0, 0, Timer - Время)
       
       UserForm1.Repaint
       'Application.ScreenUpdating = False 'включить обновление документа
       
       нашло_стр = False
       For Y = 0 To (UBound(Список_изм, 1))
         If Список_изм(Y, 0) = номер_активной_стр_из_колонт Then
            колонтитулы_всех_листов = True
            табл(a1, 1) = Список_изм(Y, 1)
            Лист = Список_изм(Y, 2)
            табл(a1, 4) = Список_изм(Y, 3)
            табл(a1, 5) = Список_изм(Y, 4)
            табл(a1, 6) = Список_изм(Y, 5)
            нашло_стр = True
               'If изм_активной_стр_из_колонт <> табл(a1, 1) Or _
               '   лист_активной_стр_из_колонт <> Лист Or _
               '   номер_докум_активной_стр_из_колонт <> табл(a1, 4) Then
                  'UserForm1.Label1.Width = CInt((номер_докум_активной_стр_из_колонт * 100) / 300)
                  'UserForm1.Label_Лист = лист_активной_стр_из_колонт & " из " & Колич_листов_в_докум
                  'UserForm1.Repaint

                  Заносим_в_колонтитул
               ' End If
                
            Exit For
         End If
       Next Y
              If (изм_активной_стр_из_колонт <> "" Or _
              лист_активной_стр_из_колонт <> "" Or _
              номер_докум_активной_стр_из_колонт <> "") And нашло_стр = False Then
                 колонтитулы_всех_листов = True
                 табл(a1, 1) = ""
                 Лист = ""
                 табл(a1, 4) = ""
                 табл(a1, 5) = ""
                 табл(a1, 6) = ""
                 табл(a1, 7) = ""
                 табл(a1, 8) = ""
                 'UserForm1.Label1.Width = CInt((номер_докум_активной_стр_из_колонт * 100) / 300)
                 'UserForm1.Label_Лист = лист_активной_стр_из_колонт & " из " & Колич_листов_в_докум
                 'UserForm1.Repaint

                 Заносим_в_колонтитул
                 колонтитулы_всех_листов = False
              End If
              нашло_стр = False
  Next i

  Application.ScreenUpdating = True 'включить обновление документа
  If (Timer - Время) > 60 Then
      Debug.Print (Timer - Время) / 60 & " мин" ' время  в сек
      Else
      Debug.Print Timer - Время & " сек" ' время  в сек
  End If
End Sub



Sub Сравнение_колонтитулов_в_листах_с_указаными_в_таблице()
Время = Timer
Application.ScreenUpdating = False 'выключить обновление документа
UserForm1.Repaint
DoEvents

'Номер таблицы с деталюхами
For i = 1 To ActiveDocument.Tables.Count Step 1
   If ActiveDocument.Tables(i).Range.Columns.Count = 18 Then
      Номер_таблицы = i
      Exit For
   End If
Next i

ActiveDocument.Tables(Номер_таблицы).Range.Cells(1).Select
номер_стр_с_первой_табл = Selection.Information(wdActiveEndPageNumber) 'номер активной стр
 
UserForm2.TextBox1.Text = ""
UserForm2.TextBox2.Text = ""
UserForm2.TextBox3.Text = ""
UserForm2.TextBox_а1_зам.Text = ""
UserForm2.TextBox_а1_нов.Text = ""
UserForm2.TextBox_а1_не.Text = ""
UserForm2.TextBox_а2_зам.Text = ""
UserForm2.TextBox_а2_нов.Text = ""
UserForm2.TextBox_а2_не.Text = ""
UserForm2.TextBox_а3_зам.Text = ""
UserForm2.TextBox_а3_нов.Text = ""
UserForm2.TextBox_а3_не.Text = ""
UserForm2.TextBox_а4_зам.Text = ""
UserForm2.TextBox_а4_нов.Text = ""
UserForm2.TextBox_а4_не.Text = ""
UserForm2.TextBox_а5_зам.Text = ""
UserForm2.TextBox_а5_нов.Text = ""
UserForm2.TextBox_а5_не.Text = ""
UserForm2.TextBox_а6_зам.Text = ""
UserForm2.TextBox_а6_нов.Text = ""
UserForm2.TextBox_а6_не.Text = ""
UserForm2.TextBox_а7_зам.Text = ""
UserForm2.TextBox_а7_нов.Text = ""
UserForm2.TextBox_а7_не.Text = ""
UserForm2.TextBox_а8_зам.Text = ""
UserForm2.TextBox_а8_нов.Text = ""
UserForm2.TextBox_а8_не.Text = ""
UserForm2.TextBox_а9_зам.Text = ""
UserForm2.TextBox_а9_нов.Text = ""
UserForm2.TextBox_а9_не.Text = ""
UserForm2.TextBox_а10_зам.Text = ""
UserForm2.TextBox_а10_нов.Text = ""
UserForm2.TextBox_а10_не.Text = ""
UserForm2.TextBox_а11_зам.Text = ""
UserForm2.TextBox_а11_нов.Text = ""
UserForm2.TextBox_а11_не.Text = ""
UserForm2.TextBox_а12_зам.Text = ""
UserForm2.TextBox_а12_нов.Text = ""
UserForm2.TextBox_а12_не.Text = ""
UserForm2.TextBox_без_изм.Text = ""
UserForm2.TextBox_не_зам.Text = ""
UserForm2.TextBox4_не_нов.Text = ""
UserForm2.TextBox_не.Text = ""
UserForm2.TextBox5.Text = ""

      If изм_текущ = "а10" And Лист_текущ = "Нов" Then UserForm2.TextBox_а10_нов.Text = UserForm2.TextBox_а10_нов.Text + Номер_стр_текущ & ", "
      If изм_текущ = "а11" And Лист_текущ = "Зам" Then UserForm2.TextBox_а11_зам.Text = UserForm2.TextBox_а11_зам.Text + Номер_стр_текущ & ", "
      If изм_текущ = "а11" And Лист_текущ = "Нов" Then UserForm2.TextBox_а11_нов.Text = UserForm2.TextBox_а11_нов.Text + Номер_стр_текущ & ", "
      If изм_текущ = "а12" And Лист_текущ = "Зам" Then UserForm2.TextBox_а12_зам.Text = UserForm2.TextBox_а12_зам.Text + Номер_стр_текущ & ", "
      If изм_текущ = "а12" And Лист_текущ = "Нов" Then UserForm2.TextBox_а12_нов.Text = UserForm2.TextBox_а12_нов.Text + Номер_стр_текущ & ", "


If Извещение_Список_текущ_изм = True Then
Else

Call Получить_список_изм_постранично 'Список_изм(1,1)

End If 'Извещение_Список_текущ_изм = True

Колич_листов_в_докум = ActiveDocument.ComputeStatistics(wdStatisticPages)






'Прогон по страницам с предпред последней
Erase Список_текущ_изм 'отчистить массив
ReDim Список_текущ_изм(Колич_листов_в_докум * 3, 5) 'задаем размер таблицы листов в 3 раза больше вдруг ктото анулирует лист это происходи крайне редко

номер_активной_стр = номер_стр_с_первой_табл - 1
  Бракованые_стр = ""
  For i = номер_стр_с_первой_табл To Колич_листов_в_докум
     DoEvents
     
     номер_активной_стр = номер_активной_стр + 1 'номер активной стр
     Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=i, Name:="" 'перейти на стр по номеру
     'Selection.Range.Select
     'Selection.HomeKey Unit:=wdLine 'перейти в начало строки

              
      'Selection.GoTo What:=wdGoToPage, Which:=wdGoToNext, Count:=1, Name:="" 'перейти на следующуу стр
      WordBasic.ViewFooterOnly ' открыть нижний колонтитул
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
      WordBasic.ViewFooterOnly ' открыть нижний колонтитул
      
       'Application.ScreenUpdating = True 'включить обновление документа
       
       'Application.ScreenUpdating = False 'включить обновление документа
       
       DoEvents
       UserForm1.Label1.Width = CInt((300 * номер_активной_стр) / Колич_листов_в_докум)
       UserForm1.Label_Лист = "Лист " & номер_активной_стр & " из " & Колич_листов_в_докум
       'If (Timer - Время) > 60 Then
          'UserForm1.Label2 = "Время " & Round((Timer - Время) / 60, 1) & " мин" ' время  в сек
       '   UserForm1.Label2 = TimeSerial(0, 0, Timer - Время)
       'Else
          'UserForm1.Label2 = "Время " & Round(Timer - Время, 1) & " сек" ' время  в сек
           UserForm1.Label2 = "Время: " & TimeSerial(0, 0, Timer - Время)
       'End If
       UserForm1.Repaint
      
      Номер_стр_текущ = Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text, Chr(13) & "", "")
      изм_текущ = Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 1).Range.Text, Chr(13) & "", "")
      Лист_текущ = Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 2).Range.Text, Chr(13) & "", "")
      Номер_докум = Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 3).Range.Text, Chr(13) & "", "")
      
      Номер_стр_текущ = Replace(Номер_стр_текущ, " ", "")
      изм_текущ = Replace(изм_текущ, " ", "")
      изм_текущ = Replace(изм_текущ, "a", "а")
      Лист_текущ = Replace(Лист_текущ, " ", "")
      Лист_текущ = Replace(Лист_текущ, ".", "")
      Лист_текущ = Replace(Лист_текущ, "зам", "Зам")
      Лист_текущ = Replace(Лист_текущ, "нов", "Нов")
      Номер_докум = Replace(Номер_докум, " ", "")
      
      
Список_текущ_изм(i, 0) = Номер_стр_текущ
Список_текущ_изм(i, 1) = изм_текущ
Список_текущ_изм(i, 2) = Лист_текущ
Список_текущ_изм(i, 3) = Номер_докум
Список_текущ_изм(i, 4) = CStr(номер_активной_стр)
      
If Извещение_Список_текущ_изм = True Then
Else
      If изм_текущ = "" Then
            UserForm2.TextBox_без_изм.Text = UserForm2.TextBox_без_изм.Text + Номер_стр_текущ & ", "
            GoTo конец_изм
      End If
      
      If изм_текущ = "а1" And Лист_текущ = "Зам" Then
            UserForm2.TextBox_а1_зам.Text = UserForm2.TextBox_а1_зам.Text + Номер_стр_текущ & ", "
            GoTo конец_изм
      End If
      If изм_текущ = "а1" And Лист_текущ = "Нов" Then
            UserForm2.TextBox_а1_нов.Text = UserForm2.TextBox_а1_нов.Text + Номер_стр_текущ & ", "
            GoTo конец_изм
      End If
      If изм_текущ = "а1" And Лист_текущ <> "Нов" And Лист_текущ <> "Зам" Then
            UserForm2.TextBox_а1_не.Text = UserForm2.TextBox_а1_не.Text + Номер_стр_текущ & ", "
            GoTo конец_изм
      End If
      If изм_текущ = "а2" And Лист_текущ = "Зам" Then
            UserForm2.TextBox_а2_зам.Text = UserForm2.TextBox_а2_зам.Text + Номер_стр_текущ & ", "
            GoTo конец_изм
      End If
      If изм_текущ = "а2" And Лист_текущ = "Нов" Then
            UserForm2.TextBox_а2_нов.Text = UserForm2.TextBox_а2_нов.Text + Номер_стр_текущ & ", "
            GoTo конец_изм
      End If
            If изм_текущ = "а2" And Лист_текущ <> "Нов" And Лист_текущ <> "Зам" Then
            UserForm2.TextBox_а2_не.Text = UserForm2.TextBox_а2_не.Text + Номер_стр_текущ & ", "
            GoTo конец_изм
      End If
      If изм_текущ = "а3" And Лист_текущ = "Зам" Then
            UserForm2.TextBox_а3_зам.Text = UserForm2.TextBox_а3_зам.Text + Номер_стр_текущ & ", "
            GoTo конец_изм
      End If
      If изм_текущ = "а3" And Лист_текущ = "Нов" Then
            UserForm2.TextBox_а3_нов.Text = UserForm2.TextBox_а3_нов.Text + Номер_стр_текущ & ", "
            GoTo конец_изм
      End If
            If изм_текущ = "а3" And Лист_текущ <> "Нов" And Лист_текущ <> "Зам" Then
            UserForm2.TextBox_а3_не.Text = UserForm2.TextBox_а3_не.Text + Номер_стр_текущ & ", "
            GoTo конец_изм
      End If
      If изм_текущ = "а4" And Лист_текущ = "Зам" Then
            UserForm2.TextBox_а4_зам.Text = UserForm2.TextBox_а4_зам.Text + Номер_стр_текущ & ", "
            GoTo конец_изм
      End If
      If изм_текущ = "а4" And Лист_текущ = "Нов" Then
            UserForm2.TextBox_а4_нов.Text = UserForm2.TextBox_а4_нов.Text + Номер_стр_текущ & ", "
            GoTo конец_изм
      End If
            If изм_текущ = "а4" And Лист_текущ <> "Нов" And Лист_текущ <> "Зам" Then
            UserForm2.TextBox_а4_не.Text = UserForm2.TextBox_а4_не.Text + Номер_стр_текущ & ", "
            GoTo конец_изм
      End If
      If изм_текущ = "а5" And Лист_текущ = "Зам" Then
            UserForm2.TextBox_а5_зам.Text = UserForm2.TextBox_а5_зам.Text + Номер_стр_текущ & ", "
            GoTo конец_изм
      End If
      If изм_текущ = "а5" And Лист_текущ = "Нов" Then
            UserForm2.TextBox_а5_нов.Text = UserForm2.TextBox_а5_нов.Text + Номер_стр_текущ & ", "
            GoTo конец_изм
      End If
            If изм_текущ = "а5" And Лист_текущ <> "Нов" And Лист_текущ <> "Зам" Then
            UserForm2.TextBox_а5_не.Text = UserForm2.TextBox_а5_не.Text + Номер_стр_текущ & ", "
            GoTo конец_изм
      End If
      If изм_текущ = "а6" And Лист_текущ = "Зам" Then
            UserForm2.TextBox_а6_зам.Text = UserForm2.TextBox_а6_зам.Text + Номер_стр_текущ & ", "
            GoTo конец_изм
      End If
      If изм_текущ = "а6" And Лист_текущ = "Нов" Then
            UserForm2.TextBox_а6_нов.Text = UserForm2.TextBox_а6_нов.Text + Номер_стр_текущ & ", "
            GoTo конец_изм
      End If
            If изм_текущ = "а6" And Лист_текущ <> "Нов" And Лист_текущ <> "Зам" Then
            UserForm2.TextBox_а6_не.Text = UserForm2.TextBox_а6_не.Text + Номер_стр_текущ & ", "
            GoTo конец_изм
      End If
      If изм_текущ = "а7" And Лист_текущ = "Зам" Then
            UserForm2.TextBox_а7_зам.Text = UserForm2.TextBox_а7_зам.Text + Номер_стр_текущ & ", "
            GoTo конец_изм
      End If
      If изм_текущ = "а7" And Лист_текущ = "Нов" Then
            UserForm2.TextBox_а7_нов.Text = UserForm2.TextBox_а7_нов.Text + Номер_стр_текущ & ", "
            GoTo конец_изм
      End If
            If изм_текущ = "а7" And Лист_текущ <> "Нов" And Лист_текущ <> "Зам" Then
            UserForm2.TextBox_а7_не.Text = UserForm2.TextBox_а7_не.Text + Номер_стр_текущ & ", "
            GoTo конец_изм
      End If
      If изм_текущ = "а8" And Лист_текущ = "Зам" Then
            UserForm2.TextBox_а8_зам.Text = UserForm2.TextBox_а8_зам.Text + Номер_стр_текущ & ", "
            GoTo конец_изм
      End If
      If изм_текущ = "а8" And Лист_текущ = "Нов" Then
            UserForm2.TextBox_а8_нов.Text = UserForm2.TextBox_а8_нов.Text + Номер_стр_текущ & ", "
            GoTo конец_изм
      End If
            If изм_текущ = "а8" And Лист_текущ <> "Нов" And Лист_текущ <> "Зам" Then
            UserForm2.TextBox_а8_не.Text = UserForm2.TextBox_а8_не.Text + Номер_стр_текущ & ", "
            GoTo конец_изм
      End If
      If изм_текущ = "а9" And Лист_текущ = "Зам" Then
            UserForm2.TextBox_а9_зам.Text = UserForm2.TextBox_а9_зам.Text + Номер_стр_текущ & ", "
            GoTo конец_изм
      End If
      If изм_текущ = "а9" And Лист_текущ = "Нов" Then
            UserForm2.TextBox_а9_нов.Text = UserForm2.TextBox_а9_нов.Text + Номер_стр_текущ & ", "
            GoTo конец_изм
      End If
            If изм_текущ = "а9" And Лист_текущ <> "Нов" And Лист_текущ <> "Зам" Then
            UserForm2.TextBox_а9_не.Text = UserForm2.TextBox_а9_не.Text + Номер_стр_текущ & ", "
            GoTo конец_изм
      End If
      If изм_текущ = "а10" And Лист_текущ = "Зам" Then
            UserForm2.TextBox_а10_зам.Text = UserForm2.TextBox_а10_зам.Text + Номер_стр_текущ & ", "
            GoTo конец_изм
      End If
      If изм_текущ = "а10" And Лист_текущ = "Нов" Then
            UserForm2.TextBox_а10_нов.Text = UserForm2.TextBox_а10_нов.Text + Номер_стр_текущ & ", "
            GoTo конец_изм
      End If
            If изм_текущ = "а10" And Лист_текущ <> "Нов" And Лист_текущ <> "Зам" Then
            UserForm2.TextBox_а10_не.Text = UserForm2.TextBox_а10_не.Text + Номер_стр_текущ & ", "
            GoTo конец_изм
      End If
      If изм_текущ = "а11" And Лист_текущ = "Зам" Then
            UserForm2.TextBox_а11_зам.Text = UserForm2.TextBox_а11_зам.Text + Номер_стр_текущ & ", "
            GoTo конец_изм
      End If
      If изм_текущ = "а11" And Лист_текущ = "Нов" Then
            UserForm2.TextBox_а11_нов.Text = UserForm2.TextBox_а11_нов.Text + Номер_стр_текущ & ", "
            GoTo конец_изм
      End If
            If изм_текущ = "а11" And Лист_текущ <> "Нов" And Лист_текущ <> "Зам" Then
            UserForm2.TextBox_а11_не.Text = UserForm2.TextBox_а11_не.Text + Номер_стр_текущ & ", "
            GoTo конец_изм
      End If
      If изм_текущ = "а12" And Лист_текущ = "Зам" Then
            UserForm2.TextBox_а12_зам.Text = UserForm2.TextBox_а12_зам.Text + Номер_стр_текущ & ", "
            GoTo конец_изм
      End If
      If изм_текущ = "а12" And Лист_текущ = "Нов" Then
            UserForm2.TextBox_а12_нов.Text = UserForm2.TextBox_а12_нов.Text + Номер_стр_текущ & ", "
            GoTo конец_изм
      End If
            If изм_текущ = "а12" And Лист_текущ <> "Нов" And Лист_текущ <> "Зам" Then
            UserForm2.TextBox_а12_не.Text = UserForm2.TextBox_а12_не.Text + Номер_стр_текущ & ", "
            GoTo конец_изм
      End If
      

      If Лист_текущ = "Зам" Then
            UserForm2.TextBox_не_зам.Text = UserForm2.TextBox_не_зам.Text + Номер_стр_текущ & ", "
            GoTo конец_изм
      End If
      If Лист_текущ = "Нов" Then
            UserForm2.TextBox4_не_нов.Text = UserForm2.TextBox4_не_нов.Text + Номер_стр_текущ & ", "
            GoTo конец_изм
      End If
      If Лист_текущ <> "Нов" And Лист_текущ <> "Зам" Then
            UserForm2.TextBox_не.Text = UserForm2.TextBox_не.Text + Номер_стр_текущ & ", "
            GoTo конец_изм
      End If
      
      UserForm2.TextBox5.Text = UserForm2.TextBox5.Text + Номер_стр_текущ & ", "
      
конец_изм:

End If 'Извещение_Список_текущ_изм = True
      
     ' 'удаляем хлам
     ' For u = 0 To UBound(Список_изм, 1) Step 1
     '   Список_изм(u, 1) = Replace(Список_изм(u, 1), "a", "а")
     '   For Y = 0 To UBound(Список_изм, 2) Step 1
     '      Список_изм(u, Y) = Replace(Список_изм(u, Y), " ", "")
     '   Next Y
     ' Next u
      
      'создаем список бракованых стр
     ' For u = 0 To UBound(Список_изм, 1) Step 1
     '    If Список_изм(u, 0) = Номер_стр_текущ Then
     '      If Список_изм(u, 1) <> Изм_текущ Or _
    '          Replace(Список_изм(u, 2), ".", "") <> Лист_текущ Or _
    '          Список_изм(u, 3) <> Номер_докум Then
    '          Бракованые_стр = Бракованые_стр & Номер_стр_текущ & ", "
    '          Сквазные_Бракованые_стр = Сквазные_Бракованые_стр & номер_активной_стр & ", "
    '          Exit For
    '       End If
    '     End If
    '  Next u
      'Список_страниц(j) = Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text, 'обращение к ячейке в табл в нижнем колонтитуле
      'Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text, Chr(13) & "", "")
      'Список_страниц(j) = Replace(Список_страниц(j), "", "")
      'Список_страниц(j) = Replace(Список_страниц(j), Chr(13), "")
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула

  Next i
  
If Извещение_Список_текущ_изм = True Then

Else
  
'Добавляем недостающие стр в Список_изм (ставим без изм)
For i = 0 To UBound(Список_текущ_изм, 1)
   If CStr(Список_текущ_изм(i, 0)) <> "" Then
     For r = 0 To UBound(Список_изм, 1)
        If CStr(Список_текущ_изм(i, 0)) = CStr(Список_изм(r, 0)) Then
           GoTo новый_номер  'если нашли
        End If
        If r = UBound(Список_изм, 1) Then 'если все просмотрели и не нашли
           For j = 0 To UBound(Список_изм, 1)
              If CStr(Список_изм(j, 0)) = "" Then
                 Список_изм(j, 0) = CStr(Список_текущ_изм(i, 0))
                 GoTo новый_номер  'если нашли
              End If
           Next j
        End If
     Next r
   
новый_номер:
    End If
Next i
'

      'удаляем хлам
      For u = 0 To UBound(Список_изм, 1) Step 1
        Список_изм(u, 1) = Replace(Список_изм(u, 1), "a", "а")
        For Y = 0 To UBound(Список_изм, 2) Step 1
           Список_изм(u, Y) = Replace(Список_изм(u, Y), " ", "")
        Next Y
      Next u
      
'Определяем не соответствие изм на листах и на  листе регистр изм. (Бракованые_стр)
For i = 0 To UBound(Список_текущ_изм, 1)
      'создаем список бракованых стр
   If CStr(Список_текущ_изм(i, 0)) <> "" Then
      For u = 0 To UBound(Список_изм, 1) Step 1
         If CStr(Список_изм(u, 0)) = CStr(Список_текущ_изм(i, 0)) Then
           If CStr(Список_изм(u, 1)) <> CStr(Список_текущ_изм(i, 1)) Or _
              Replace(Список_изм(u, 2), ".", "") <> Список_текущ_изм(i, 2) Or _
              CStr(Список_изм(u, 3)) <> CStr(Список_текущ_изм(i, 3)) Then
              If CStr(Список_изм(u, 1)) = "" Then
                изм = "без"
              Else
                изм = CStr(Список_изм(u, 1))
              End If
              If CStr(Список_текущ_изм(i, 1)) = "" Then
                изм_текущ = "без"
              Else
                изм_текущ = CStr(Список_текущ_изм(i, 1))
              End If
              Бракованые_стр = Бракованые_стр & Список_текущ_изм(i, 0) & "-" & изм & "/" & изм_текущ & ", "
              Сквазные_Бракованые_стр = Сквазные_Бракованые_стр & Список_текущ_изм(i, 4) & ", "
              Exit For
           End If
         End If
      Next u
  End If
Next i
  
If Бракованые_стр <> "" Then
  UserForm2.TextBox1.Text = Mid(Бракованые_стр, 1, Len(Бракованые_стр) - 2)
  UserForm2.TextBox2.Text = Mid(Сквазные_Бракованые_стр, 1, Len(Сквазные_Бракованые_стр) - 2)
End If


  Application.ScreenUpdating = True 'включить обновление документа
  UserForm1.Hide

  If Бракованые_стр = "" Then
    сообщение = MsgBox("Бракованых колонтитулов не найдено", vbOKOnly + vbInformation, "Проверка колонтитулов")
  Else
    UserForm2.Show
  End If
End If 'Извещение_Список_текущ_изм = True

  If (Timer - Время) > 60 Then
      Debug.Print (Timer - Время) / 60 & " мин" ' время  в сек
      Else
      Debug.Print Timer - Время & " сек" ' время  в сек
  End If


End Sub



Sub Удаляем_разрывы()

Application.ScreenUpdating = False 'выключить обновление документа
Dim Список_новых_листов_с_разрывами(), Список_отчисченых_стр()

Колич_листов_в_докум = ActiveDocument.ComputeStatistics(wdStatisticPages)
Колич_разделоа_в_докум = ActiveDocument.Sections.Count 'количество разделов
'номер_активной_стр = Selection.Information(wdActiveEndPageNumber) 'номер активной стр
'Selection.Information (wdActiveEndSectionNumber) 'номер активного раздела
'ActiveDocument.Sections.Count'количество разделов
'Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=Колич_листов_в_докум - 2, Name:="" 'перейти на стр по номеру
'Selection.GoTo What:=wdGoToSection, Which:=wdGoToFirst, Count:=5, Name:="" 'перейти в раздел

'Получть список страниц
    сообщение = MsgBox("Получить список страниц из файла TXT: Yes-""Да"", No-""Создать новый список"", Cancel-""Выйти"" ", vbYesNoCancel + vbQuestion + vbDefaultButton1, "Запрос на список страниц")
    If сообщение = vbCancel Then Exit Sub
    'If сообщение = vbYes Then Лист = "Зам."
    If сообщение = vbNo Then Call ЗАГРУЗКА_Список_стр_в_файл


путь = ActiveDocument.Path
имя = ActiveDocument.Name

'проверка существует ли файл
If Dir(путь & "\" & имя & "_страницы.txt") = "" Then
  сообщение = MsgBox("Файл не найден. Создать новый OK-""Да"",Cancel-""Выйти"" ", vbOKCancel + vbQuestion + vbDefaultButton1, "Запрос")
  If сообщение = vbCancel Then Exit Sub
  If сообщение = vbOK Then Call ЗАГРУЗКА_Список_стр_в_файл
End If

Open путь & "\" & имя & "_страницы.txt" For Input As #1  'открыть для чтения "Input"
Line Input #1, текст_из_txt 'чтение строки
Close #1 'закрыть документ

Массив = Split(текст_из_txt, ",") 'получаем массив
ReDim Список_страниц(UBound(Массив, 1)) 'задаем размер таблицы
ReDim Список_новых_листов_с_разрывами(UBound(Список_страниц, 1), UBound(Список_страниц, 1)) 'задаем размер таблицы
ReDim Список_отчисченых_стр(Колич_листов_в_докум, 2)



'присваеваем значения оного массива другому
For i = 0 To UBound(Массив, 1)
   Список_страниц(i) = Массив(i)
Next i

номер = 0

'Прогон по страницам с пред последней
  For i = 0 To UBound(Список_страниц, 1) Step 1
      'добавить предпоследнию стр в список
      If (i = UBound(Список_страниц, 1) - 1) And CInt(Список_страниц(UBound(Список_страниц, 1) - 1 - 1)) = Список_страниц(UBound(Список_страниц, 1) - 1 - 1) Then
        Список_новых_листов_с_разрывами(номер, 0) = Список_страниц(i + 1) 'номер не целых страниц
        Список_новых_листов_с_разрывами(номер, 1) = CStr(i + 1 + 1) 'номер страницы сквазная
        номер = номер + 1
      End If
     'Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=i, Name:="" 'перейти на стр по номеру
      'добавить не целые стр в писок
      If CInt(Список_страниц(i)) <> Список_страниц(i) Then
        'ReDim Preserve Список_новых_листов_с_разрывами(UBound(Список_новых_листов_с_разрывами, 1) + 1)

        Список_новых_листов_с_разрывами(номер, 0) = Список_страниц(i) 'номер не целых страниц
        Список_новых_листов_с_разрывами(номер, 1) = CStr(i + 1) 'номер страницы сквазная
        номер = номер + 1
      End If
  Next i

'определение номера активной таблицы
'ActiveDocument.Range.Tables.Count


'Set tblSel = Selection.Tables(1)
'ingStart = tblSel.Range.Start
'Номер таблицы с деталюхами
For i = 1 To ActiveDocument.Tables.Count Step 1
   If ActiveDocument.Tables(i).Range.Columns.Count = 18 Then
      Номер_таблицы = i
      Exit For
   End If
Next i

ActiveDocument.Tables(Номер_таблицы).Range.Cells(1).Select
номер_стр_с_первой_табл = Selection.Information(wdActiveEndPageNumber) 'номер активной стр

For i = UBound(Список_новых_листов_с_разрывами, 1) To 0 Step -1
   If Список_новых_листов_с_разрывами(i, 1) <> Empty Then
     If i <> 0 Then
       If CInt(Список_новых_листов_с_разрывами(i, 0)) <> (Список_новых_листов_с_разрывами(i, 0)) And _
          CInt(Список_новых_листов_с_разрывами(i - 1, 0)) <> (Список_новых_листов_с_разрывами(i - 1, 0)) And _
          (Список_новых_листов_с_разрывами(i, 0)) - (Список_новых_листов_с_разрывами(i - 1, 0)) <= 0.1 Then
          GoTo Закончить_цикл
       End If
     End If
      конечная_стр = Список_новых_листов_с_разрывами(i, 1) - 1
      If i = 0 Then
              начальная_стр = номер_стр_с_первой_табл
        Else: начальная_стр = Список_новых_листов_с_разрывами(i - 1, 1) + 1
      End If
      'If конечная_стр - начальная_стр = 1 Then GoTo Закончить_цикл
      
     
      
      Номера_стр_конеч = Список_страниц(конечная_стр - 1)
      Номера_стр_начал = Список_страниц(начальная_стр - 1)
            
     ' If (CInt(Номера_стр_начал) = Номера_стр_начал) And _
     '    (CInt(Номера_стр_конеч) <> Номера_стр_конеч) And _
     '    Номера_стр_начал - Номера_стр_конеч < 1 Then
     ' GoTo Закончить_цикл
     '  End If
          
      'End If
      'If CInt(Список_страниц(конечная_стр - 2)) = (Список_страниц(конечная_стр - 2)) Then
      '    Нужный_номер_стр = конечная_стр - 1
      'End If
      
      
      Нужный_номер_стр = конечная_стр
      Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=Нужный_номер_стр, Name:="" 'перейти на стр по номеру
      
       DoEvents
       UserForm1.Label1.Width = CInt((300 * (Колич_листов_в_докум - номер_активной_стр)) / Колич_листов_в_докум)
       UserForm1.Label_Лист = "Лист " & (Колич_листов_в_докум - номер_активной_стр) & " из " & Колич_листов_в_докум
       UserForm1.Repaint
      
      
      For Y = конечная_стр To (начальная_стр + 1) Step -1
         'If Y = (начальная_стр + 1) Then Exit For
    'ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader 'перейти в верхний колонтитул
    'Selection.HeaderFooter.LinkToPrevious = Not Selection.HeaderFooter.LinkToPrevious ' отключить как в предудущем разделе
    'ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument 'Закрыть колонтитул
         If Selection.Information(wdActiveEndSectionNumber) = 3 Then  'номер активного раздела
            ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader 'перейти в верхний колонтитул
            Selection.HeaderFooter.LinkToPrevious = False ' отключить как в предудущем разделе
            ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument 'Закрыть колонтитул
            Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=Нужный_номер_стр, Name:="" 'перейти на стр по номеру
         End If
         Selection.Find.ClearFormatting
         'Selection.Find.Replacement.ClearFormatting
         With Selection.Find
              .Text = "^b"
              .Replacement.Text = ""
                 .Forward = False
              .Wrap = wdFindAsk
              .Format = False
              .MatchCase = False
              .MatchWholeWord = False
              .MatchWildcards = False
              .MatchSoundsLike = False
              .MatchAllWordForms = False
         End With
         Selection.Find.Execute
                 
          WordBasic.ViewFooterOnly ' открыть нижний колонтитул
          
          On Error GoTo Закончить_цикл 'пропускаем ошибки
          Номер_текущ_стр = Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text, Chr(13) & "", "") 'обращение к ячейке в табл в нижнем колонтитуле
          On Error GoTo 0
          
            For h = 0 To UBound(Список_страниц, 1) Step 1
                If Номер_текущ_стр = Список_страниц(h) Then
                   номер_стр = h + 1
                End If
            Next h
          
          ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула

         If номер_стр < начальная_стр Then
            GoTo Закончить_цикл
         End If
         Selection.Delete Unit:=wdCharacter, Count:=1
         
         Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=Нужный_номер_стр, Name:="" 'перейти на стр по номеру

      
         'Проверка колонтитула
          WordBasic.ViewFooterOnly ' открыть нижний колонтитул
          ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
          WordBasic.ViewFooterOnly ' открыть нижний колонтитул
          Номер_текущ_стр = Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text, Chr(13) & "", "") 'обращение к ячейке в табл в нижнем колонтитуле
          
          'чистака колонтитула
          Selection.HeaderFooter.Range.Tables(1).Cell(2, 1).Range.Text = "" 'таблица в колонтитуле
          Selection.HeaderFooter.Range.Tables(1).Cell(2, 2).Range.Text = "" 'таблица в колонтитуле
          Selection.HeaderFooter.Range.Tables(1).Cell(2, 3).Range.Text = "" 'таблица в колонтитуле
          Selection.HeaderFooter.Range.Tables(1).Cell(2, 4).Range.Text = "" 'таблица в колонтитуле
          Selection.HeaderFooter.Range.Tables(1).Cell(2, 5).Range.Text = "" 'таблица в колонтитуле
          'Список_отчисченых_стр()
          
          'номер_активной_стр = Y 'Selection.Information(wdActiveEndPageNumber) 'номер активной стр
        
       номер_активной_стр = Selection.Information(wdActiveEndPageNumber) 'номер активной стр
         
         If номер_активной_стр < номер_стр_с_первой_табл Then
            GoTo Закончить_цикл
         End If
         
       DoEvents
       UserForm1.Label1.Width = CInt((300 * (Колич_листов_в_докум - номер_активной_стр)) / Колич_листов_в_докум)
       UserForm1.Label_Лист = "Лист " & (Колич_листов_в_докум - номер_активной_стр) & " из " & Колич_листов_в_докум
       UserForm1.Repaint


 
 
       'If Номер_текущ_стр <> Список_страниц(Нужный_номер_стр - 1) Then '(Selection.Information(wdActiveEndPageNumber)) Then
       If Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text, Chr(13) & "", "") <> Список_страниц(Selection.Information(wdActiveEndPageNumber) - 1) And _
        Список_страниц(Selection.Information(wdActiveEndPageNumber) - 1) <> 1 Then
            'Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=номер_активной_стр + 1, Name:="" 'перейти на стр по номеру
           'WordBasic.ViewFooterOnly ' открыть нижний колонтитул
            If Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text, Chr(13) & "", "") <> Список_страниц(Selection.Information(wdActiveEndPageNumber) - 1) Then
               With Selection.HeaderFooter.PageNumbers  ' задать номер стр номером
                 .NumberStyle = wdPageNumberStyleArabic
                 .HeadingLevelForChapter = 0
                 .IncludeChapterNumber = False
                 .ChapterPageSeparator = wdSeparatorHyphen
                 .RestartNumberingAtSection = True
                 .StartingNumber = Список_страниц(начальная_стр - 1)
              End With
            End If
           
           'Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=Нужный_номер_стр, Name:="" 'перейти на стр по номеру
           'WordBasic.ViewFooterOnly ' открыть нижний колонтитул
           
           If Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text, Chr(13) & "", "") <> Список_страниц(Selection.Information(wdActiveEndPageNumber) - 1) Then
              With Selection.HeaderFooter.PageNumbers  ' задать номер стр номером
                 .NumberStyle = wdPageNumberStyleArabic
                 .HeadingLevelForChapter = 0
                 .IncludeChapterNumber = False
                 .ChapterPageSeparator = wdSeparatorHyphen
                 .RestartNumberingAtSection = True
                 .StartingNumber = 1
              End With
           End If

            If Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text, Chr(13) & "", "") <> Список_страниц(Selection.Information(wdActiveEndPageNumber) - 1) Then
             With Selection.HeaderFooter.PageNumbers  ' продолжить нумацию
               .NumberStyle = wdPageNumberStyleArabic
               .HeadingLevelForChapter = 0
               .IncludeChapterNumber = False
               .ChapterPageSeparator = wdSeparatorHyphen
               .RestartNumberingAtSection = False
               .StartingNumber = 0
             End With
            End If
      End If
           ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
           
           Нужный_номер_стр = Нужный_номер_стр - 1
         
      Next Y
   End If
Закончить_цикл:
ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула

Next i

'заменить разрыв на "со следующей страницы"
Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=номер_стр_с_первой_табл, Name:="" 'перейти на стр по номеру

 With Selection.PageSetup
        .LineNumbering.Active = False
        '.Orientation = wdOrientLandscape
        '.TopMargin = CentimetersToPoints(-6.1)
        '.BottomMargin = CentimetersToPoints(-3)
        '.LeftMargin = CentimetersToPoints(2.05)
        '.RightMargin = CentimetersToPoints(1.73)
        '.Gutter = CentimetersToPoints(1)
        '.HeaderDistance = CentimetersToPoints(0.96)
        '.FooterDistance = CentimetersToPoints(0)
        '.PageWidth = CentimetersToPoints(42)
        '.PageHeight = CentimetersToPoints(29.7)
        .FirstPageTray = wdPrinterDefaultBin
        .OtherPagesTray = wdPrinterDefaultBin
        .SectionStart = wdSectionNewPage
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .VerticalAlignment = wdAlignVerticalTop
        .SuppressEndnotes = False
        .MirrorMargins = False
        .TwoPagesOnOne = False
        .BookFoldPrinting = False
        .BookFoldRevPrinting = False
        .BookFoldPrintingSheets = 1
        .GutterPos = wdGutterPosLeft
 End With

    'ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader 'перейти в верхний колонтитул
    'Selection.HeaderFooter.LinkToPrevious = Not Selection.HeaderFooter.LinkToPrevious ' отключить как в предудущем разделе
    'ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument 'Закрыть колонтитул
Application.ScreenUpdating = True 'выключить обновление документа

End Sub






Sub ЗАГРУЗКА_Проверка_и_замена_всех_колонт()
  'UserForm1.Repaint
  UserForm1.Label1.Width = 0
  'DoEvents
  ГАЛКА_Проверка_и_замена_всех_колонт = True
  UserForm1.Show
  ГАЛКА_Проверка_и_замена_всех_колонт = False
  колонтитулы_всех_листов = False
  'UserForm1.Hide
End Sub


Sub ЗАГРУЗКА_Список_стр_в_файл()
  'UserForm1.Repaint
   UserForm1.Label1.Width = 0
   UserForm1.Label_Лист = "Подготовка"
  'DoEvents
  ГАЛКА_Список_стр_в_файл = True
  On Error Resume Next
  UserForm1.Show
  On Error GoTo 0  'сново не пропускаем ошибки
  ГАЛКА_Список_стр_в_файл = False
  'UserForm1.Hide
End Sub

Sub ЗАГРУЗКА_Сравнение_колонтитулов_в_листах_с_указаными_в_таблице()
    'сообщение = MsgBox("Выполнить поиск или просто открыть окно с результатами прошлого поиска: Yes-""Поиск"", No-""Открыть Форму"", Cancel-""Выйти"" ", vbYesNoCancel + vbQuestion + vbDefaultButton1, "Запрос на список страниц")
    'If сообщение = vbCancel Then Exit Sub
    'If сообщение = vbYes Then Лист = "Зам."
    'If сообщение = vbNo Then
    '  UserForm2.Show
    '  Exit Sub
    'End If
  'UserForm1.Repaint
   UserForm1.Label1.Width = 0
   UserForm1.Label_Лист = "Подготовка"
  'DoEvents
  ГАЛКА_Сравнение_колонтитулов_в_листах_с_указаными_в_таблице = True
  'не_меняем_размер_формы = True
  UserForm1.Show
  ГАЛКА_Сравнение_колонтитулов_в_листах_с_указаными_в_таблице = False
  'UserForm1.Hide
End Sub

Sub ЗАГРУЗКА_Удаляем_разрывы()
  'UserForm1.Repaint
   UserForm1.Label1.Width = 0
   UserForm1.Label_Лист = "Подготовка"
  'DoEvents
  ГАЛКА_Удаляем_разрывы = True
  UserForm1.Show
  ГАЛКА_Удаляем_разрывы = False
  'UserForm1.Hide
End Sub
'Удаляем_разрывы

Sub Инструкция()
  UserForm3.Show
End Sub



Sub Красим_в_Зеленый()
Attribute Красим_в_Зеленый.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Красим_в_Зеленый"
'
' Красим_в_Зеленый Макрос
'
'
    If Selection.Range.HighlightColorIndex = wdBrightGreen Then
       Selection.Range.HighlightColorIndex = wdNoHighlight
    Else
       Selection.Range.HighlightColorIndex = wdBrightGreen
    End If
End Sub
Sub Красим_в_Красный()
Attribute Красим_в_Красный.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Красим_в_Красный"
'
' Красим_в_Красный Макрос
'
'
    If Selection.Range.HighlightColorIndex = wdRed Then
       Selection.Range.HighlightColorIndex = wdNoHighlight
    Else
       Selection.Range.HighlightColorIndex = wdRed
    End If
    'Options.DefaultHighlightColorIndex = wdRed
End Sub
Sub Красим_в_Желтый_1()
Attribute Красим_в_Желтый_1.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Красим_в_Желтый"
'
' Красим_в_Желтый_1 Макрос
'
'
    If Selection.Range.HighlightColorIndex = wdYellow Then
       Selection.Range.HighlightColorIndex = wdNoHighlight
    Else
       Selection.Range.HighlightColorIndex = wdYellow
    End If
    'Selection.Range.HighlightColorIndex = wdYellow
End Sub

Sub Красим_в_Голубой()
'
' Красим_в_Голубой Макрос
'
'
    If Selection.Range.HighlightColorIndex = wdTurquoise Then
       Selection.Range.HighlightColorIndex = wdNoHighlight
    Else
       Selection.Range.HighlightColorIndex = wdTurquoise
    End If
    'Selection.Range.HighlightColorIndex = wdTurquoise
End Sub

Sub Красим_в_Серый()
'
' Красим_в_Серый Макрос
'
'
    If Selection.Range.HighlightColorIndex = wdGray25 Then
       Selection.Range.HighlightColorIndex = wdNoHighlight
    Else
       Selection.Range.HighlightColorIndex = wdGray25
    End If
End Sub

Sub ЗАГРУЗКА_Замена_точек_на_запятые()
  'UserForm1.Repaint
   UserForm1.Label1.Width = 0
   UserForm1.Label_Лист = "Подготовка"
  'DoEvents
  ГАЛКА_Замена_точек_на_запятые = True
  UserForm1.Show
  ГАЛКА_Замена_точек_на_запятые = False
  'UserForm1.Hide
End Sub





Sub Удалить_пустые_строки()

'определение номера активной таблицы
Set tblSel = Selection.Tables(1)
ingStart = tblSel.Range.Start
For i = 1 To ActiveDocument.Tables.Count Step 1
 If ActiveDocument.Tables(i).Range.Start = ingStart Then
    ingTbIndex = i
    Exit For
 End If
 Next i
 
'Удаляем строку если 2 и 4 ячейки пуста
For i = 1 To ActiveDocument.Tables(ingTbIndex).Rows.Count Step 1 'прогон по строкам
   'ActiveDocument.Tables(1).Rows.Select
   'For r = 1 To ActiveDocument.Tables(у1).Columns.Count Step 1
   'If Err.Number = 5941 Then GoTo переход
   'On Error GoTo переход  'пропускаем ошибки
    On Error Resume Next
    f = ActiveDocument.Tables(ingTbIndex).Cell(i, 2).Range.Text 'Если заглючит Err.Number подхватит ошибку
    If Err.Number = 5941 Then
      GoTo переход
    End If
    f = ActiveDocument.Tables(ingTbIndex).Cell(i, 4).Range.Text 'Если заглючит Err.Number подхватит ошибку
    If Err.Number = 5941 Then
      GoTo переход
    End If
        If Replace(ActiveDocument.Tables(ingTbIndex).Cell(i, 2).Range.Text, Chr(13) & "", "") = "" And Replace(ActiveDocument.Tables(ingTbIndex).Cell(i, 4).Range.Text, Chr(13) & "", "") = "" And Err.Number <> 5941 Then
          'On Error GoTo переход  'пропускаем ошибки
          ActiveDocument.Tables(ingTbIndex).Cell(i, 1).Select
          'ActiveDocument.Tables(ingTbIndex).Rows(i).Select
          Selection.Rows.Delete
        End If
переход:
    Err.Clear

    DoEvents
Next i
On Error GoTo 0
End Sub

Sub Замена_точек_на_запятые()
  
t = Timer

Application.ScreenUpdating = False 'отключить обновление документа
количество_таблиц = ActiveDocument.Tables.Count

For i = 1 To ActiveDocument.Tables.Count Step 1
       'номер_активной_стр = Selection.Information(wdActiveEndPageNumber) 'номер активной стр
       количество_строк = ActiveDocument.Tables(i).Range.Rows.Count
       DoEvents
       UserForm1.Label1.Width = CInt((300 * i) / количество_таблиц)
       UserForm1.Label_Лист = "Таблица " & i & " из " & количество_таблиц & " (строка " & g & " из " & количество_строк & " )"
       'If (Timer - t) > 60 Then
       '   UserForm1.Label2 = "Время " & Round((Timer - t) / 60, 1) & " мин" ' время  в сек
       'Else
       '   UserForm1.Label2 = "Время " & Round(Timer - t, 1) & " сек" ' время  в сек
       'End If
       UserForm1.Label2 = "Время: " & TimeSerial(0, 0, Timer - t)
       
       UserForm1.Repaint
 
 If ActiveDocument.Tables(i).Range.Columns.Count = 18 Then
    For g = 1 To ActiveDocument.Tables(i).Range.Rows.Count Step 1
        
        DoEvents
        UserForm1.Label2 = "Время: " & TimeSerial(0, 0, Timer - t)
        UserForm1.Label_Лист = "Таблица " & i & " из " & количество_таблиц & " (строка " & g & " из " & количество_строк & " )"
        UserForm1.Repaint
        
        If ActiveDocument.Tables(i).Cell(g, 5).Range <> Chr(13) & "" Or _
           ActiveDocument.Tables(i).Cell(g, 6).Range <> Chr(13) & "" Or _
           ActiveDocument.Tables(i).Cell(g, 7).Range <> Chr(13) & "" Then
           'Replace(МассаЕд, Chr(13) & "", "")
           код = ActiveDocument.Tables(i).Cell(g, 5).Range
           колич = ActiveDocument.Tables(i).Cell(g, 6).Range
           общ = ActiveDocument.Tables(i).Cell(g, 7).Range
           код = Replace(код, ".", ",")
           колич = Replace(колич, ".", ",")
           общ = Replace(общ, ".", ",")
           код = Replace(код, Chr(13) & "", "")
           колич = Replace(колич, Chr(13) & "", "")
           общ = Replace(общ, Chr(13) & "", "")
           код = Replace(код, " ", "")
           колич = Replace(колич, " ", "")
           общ = Replace(общ, " ", "")
        
           ActiveDocument.Tables(i).Cell(g, 5).Range = код
           ActiveDocument.Tables(i).Cell(g, 6).Range = колич
           ActiveDocument.Tables(i).Cell(g, 7).Range = общ
        End If
    Next g
 End If
 Next i
 
     If (Timer - t) > 60 Then
      Debug.Print (Timer - t) / 60 & " мин" ' время  в сек
      Else
      Debug.Print Timer - t & " сек" ' время  в сек
    End If
    
 
Application.ScreenUpdating = True 'включить обновление документа

End Sub


Sub Вносим_в_колонтитулы_только_последние_изменения_тест()


Sub Вносим_в_колонтитулы_только_последние_изменения()

 Dim список_стр_изм()
 
  Время = Timer
  
  DoEvents
  
'Номер таблицы с деталюхами
For i = 1 To ActiveDocument.Tables.Count Step 1
   If ActiveDocument.Tables(i).Range.Columns.Count = 18 Then
      Номер_таблицы = i
      Exit For
   End If
Next i

ActiveDocument.Tables(Номер_таблицы).Range.Cells(1).Select
номер_стр_с_первой_табл = Selection.Information(wdActiveEndPageNumber) 'номер активной стр

Application.ScreenUpdating = False 'отключить обновление документа
  'UserForm1.Label1.Width = 0
  'UserForm1.Show
Колич_листов_в_докум = ActiveDocument.ComputeStatistics(wdStatisticPages)

   'Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=номер_активной_стр, Name:="" 'перейти на стр по номеру


'Получаем данные о последнем измении из последней табл
Call Получить_список_изм_постранично  'Список_изм()



'Erase табл 'отчистить массив
'ReDim табл(.Rows.Count * 2, .Rows.Count * 2) 'задаем размер таблицы
'For i = 0 To UBound(Список_изм, 1)
'     табл(i, 1) = Список_изм(i, 1)
'     табл(i, 2) = Список_изм(i, 2)
'     табл(i, 4) = Список_изм(i, 3)
'     табл(i, 5) = Список_изм(i, 4)
'     табл(i, 6) = Список_изм(i, 5)
'Next i

'Определить номер последнего измения
   With ActiveDocument.Range.Tables(ActiveDocument.Range.Tables.Count)
      For i = 4 To .Rows.Count  ' получение стр из последней табл для каждого изм (а1, а2 ...)
         If .Cell(i, 1).Range.Text <> Chr(13) & "" Then
            изм_последн = Replace(.Cell(i, 1).Range.Text, Chr(13) & "", "")
            'Лист = "Зам."

         End If
      Next i
   End With
   
   
   
'Заносим последнее изм. в колонт постранично
For i = номер_стр_с_первой_табл To Колич_листов_в_докум - 1



       Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=i, Name:="" 'перейти на стр по номеру
       

       номер_активной_стр = Selection.Information(wdActiveEndPageNumber) 'номер активной стр
       DoEvents
       UserForm1.Label1.Width = CInt((300 * номер_активной_стр) / Колич_листов_в_докум)
       UserForm1.Label_Лист = "Лист " & номер_активной_стр & " из " & Колич_листов_в_докум
       'If (Timer - t) > 60 Then
       '   UserForm1.Label2 = "Время " & Round((Timer - t) / 60, 1) & " мин" ' время  в сек
       'Else
       '   UserForm1.Label2 = "Время " & Round(Timer - t, 1) & " сек" ' время  в сек
       'End If
       UserForm1.Label2 = "Время: " & TimeSerial(0, 0, Timer - Время)
       UserForm1.Repaint
       
       'For g = 0 To UBound(Список_изм, 1)

       
       WordBasic.ViewFooterOnly ' открыть нижний колонтитул
       ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
       WordBasic.ViewFooterOnly ' открыть нижний колонтитул
       
       'Поиск есть ли на текущей стр последнее изм.
       For g = 0 To UBound(Список_изм, 1)
           If CStr(Список_изм(g, 1)) = CStr(изм_последн) And CStr(Список_изм(g, 0)) = Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text, Chr(13) & "", "") Then
              
              Лист = Список_изм(g, 2)
              колонтитулы_листов_по_последн_изм = True
              Call Заносим_в_колонтитул
              колонтитулы_листов_по_последн_изм = False
              
       '       Selection.HeaderFooter.Range.Tables(1).Cell(2, 1).Range.Text = Список_изм(g, 1)
       '       Selection.HeaderFooter.Range.Tables(1).Cell(2, 2).Range.Text = Список_изм(g, 2)
       '       Selection.HeaderFooter.Range.Tables(1).Cell(2, 3).Range.Text = Список_изм(g, 3)
       '       Selection.HeaderFooter.Range.Tables(1).Cell(2, 4).Range.Text = Список_изм(g, 4)
       '       Selection.HeaderFooter.Range.Tables(1).Cell(2, 5).Range.Text = Список_изм(g, 5)
       '       ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
              Exit For
             'номер_активной_стр_из_колонт = Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text, Chr(13) & "", "")
             'изм_активной_стр_из_колонт = Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 1).Range.Text, Chr(13) & "", "")
              'лист_активной_стр_из_колонт = Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 2).Range.Text, Chr(13) & "", "")
             'номер_докум_активной_стр_из_колонт = Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 3).Range.Text, Chr(13) & "", "")
           End If
       Next g

    

Next i

  ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
  
  Application.ScreenUpdating = True 'включить обновление документа
  If (Timer - Время) > 60 Then
      Debug.Print (Timer - Время) / 60 & " мин" ' время  в сек
      Else
      Debug.Print Timer - Время & " сек" ' время  в сек
  End If
  
End Sub


Sub ЗАГРУЗКА_Вносим_в_колонтитулы_только_последние_изменения()
  'UserForm1.Repaint
   UserForm1.Label1.Width = 0
   UserForm1.Label_Лист = "Подготовка"
  'DoEvents
  ГАЛКА_Вносим_в_колонтитулы_только_последние_изменения = True
  UserForm1.Show
  ГАЛКА_Вносим_в_колонтитулы_только_последние_изменения = False
  'UserForm1.Hide
End Sub

Sub ЗАГРУЗКА_Пересчет_всех_пунктов()
  'UserForm1.Repaint
   UserForm1.Label1.Width = 0
   UserForm1.Label_Лист = "Подготовка"
  'DoEvents
  ГАЛКА_Пересчет_всех_пунктов = True
  UserForm1.Show
  ГАЛКА_Пересчет_всех_пунктов = False
  'UserForm1.Hide
End Sub


Sub А_Сокращения()
Dim r As Range, x, cl As New Collection, s$
  Set r = ActiveDocument.Range
  On Error Resume Next
  With r.Find
    .Text = "<[А-ЯЁ]{2;}>" 'ГОСТ [0-9.-]{2;} <[А-ЯЁ]{2;}>
    .Forward = True
    .Wrap = wdFindStop
    .Format = False
    .MatchWildcards = True
    While .Execute
      s = r.Text
      For x = 1 To cl.Count
        If s < cl(x) Then cl.Add s, s, Before:=x: GoTo 1
      Next
      cl.Add s, s
1   Wend
  End With
  With ActiveDocument.Range
    .InsertParagraphAfter
    For Each x In cl
      .InsertAfter vbCr & x
    Next
  End With
End Sub


Sub Пересчет_всех_пунктов()
t = Timer
Колич_листов_в_докум = ActiveDocument.ComputeStatistics(wdStatisticPages)

For i = 1 To ActiveDocument.Tables.Count Step 1 'Прогон по таблицам
  For j = 1 To ActiveDocument.Tables(i).Rows.Count  'Прогон по строкам
     If ActiveDocument.Tables(i).Columns.Count = 18 Then
        ActiveDocument.Tables(i).Cell(j, 2).Select
        
       DoEvents
       номер_активной_стр = Selection.Information(wdActiveEndPageNumber) 'номер активной стр       DoEvents
       UserForm1.Label1.Width = CInt((300 * номер_активной_стр) / Колич_листов_в_докум)
       UserForm1.Label_Лист = "Лист " & номер_активной_стр & " из " & Колич_листов_в_докум
       'If (Timer - t) > 60 Then
       '   UserForm1.Label2 = "Время " & Round((Timer - t) / 60, 1) & " мин" ' время  в сек
       'Else
       '   UserForm1.Label2 = "Время " & Round(Timer - t, 1) & " сек" ' время  в сек
       'End If
       UserForm1.Label2 = "Время: " & TimeSerial(0, 0, Timer - t)
       
       UserForm1.Repaint
       
        Call Сумм_Масса
     End If
  Next j
Next i
End Sub


Sub Канкулятор_в_ячейке()
Dim objExcApp As Object
'у1 = Selection.Rows.First.Index  'номер активной строки в таблице
'х1 = Selection.Columns.First.Index  'номер активного столбца в таблице
 у1 = Selection.Information(wdEndOfRangeRowNumber) 'номер строки ячейки таблицы к которой тыкнут курсор
 х1 = Selection.Information(wdEndOfRangeColumnNumber) 'номер столбца ячейки таблицы к которой тыкнут курсор


'определение номера активной таблицы
Set tblSel = Selection.Tables(1)
ingStart = tblSel.Range.Start
For i = 1 To ActiveDocument.Tables.Count Step 1
 If ActiveDocument.Tables(i).Range.Start = ingStart Then
    ingTbIndex = i
    Exit For
 End If
 Next i
 
текст_яч = ActiveDocument.Tables(ingTbIndex).Cell(у1, х1).Range.Text    'снимаем значение с ячейки в таблице
текст_яч = Replace(текст_яч, Chr(13) & "", "")
текст_яч = Replace(текст_яч, ",", ".")

'Работа с экселем

On Error Resume Next 'пропускаем ошибки
Set objExcApp = GetObject(, "Excel.Application")
  'Если эксель не открыт
  If objExcApp Is Nothing Then
     On Error GoTo 0  'сново не пропускаем ошибки
     Set objExcApp = CreateObject("Excel.Application")
'     objExcApp.Workbooks.Add 'создать книгу эксель
     'objExcApp.Visible = True 'сделать эксель видемым
     'заносим данные в ячейку таблицы ворд
     On Error Resume Next
     Результат = objExcApp.Application.Evaluate(текст_яч)
     Результат = Replace(Результат, ".", ",")
     ActiveDocument.Tables(ingTbIndex).Cell(у1, х1).Range = Результат
     On Error GoTo 0  'сново не пропускаем ошибки
     objExcApp.Quit 'закрыть программы эксель
     GoTo В_Конец
  'Если эксель Открыт
  Else
     On Error Resume Next
     Результат = objExcApp.Application.Evaluate(текст_яч)
     Результат = Replace(Результат, ".", ",")
     ActiveDocument.Tables(ingTbIndex).Cell(у1, х1).Range = Результат
     On Error GoTo 0  'сново не пропускаем ошибки
'     Set objExcDoc = objExcApp.Workbooks.Application
'     objExcDoc.Sheets.Add After:=objExcDoc.Sheets(objExcDoc.Sheets.Count) 'создать лист в конце экселя
     'objExcApp.Visible = True 'сделать эксель видемым
  End If
  
В_Конец:

Call Сумм_Масса

End Sub



Sub ГОСТ_Обновить_год()
Dim objExcApp As Object
'Dim ie As Object


Dim Shell_Object

If x_ГОСТ = СписокГОСТов.Count Or x_ГОСТ = 0 Then
'Set ie = CreateObject("InternetExplorer.Application")
Set objshell = CreateObject("Wscript.shell")
Set ie = New InternetExplorerMedium
End If

'Запрос = "ГОСТ 19903-"
If ПОЛНЫЙ_поиск = True Then
  ГОСТ_целеком = Искомый_ГОСТ
Else
  ГОСТ_целеком = Selection.Text
End If

ГОСТ_целеком = Replace(ГОСТ_целеком, Chr(13) & "", "")

Массив = Split(ГОСТ_целеком, "-")
колич_элем = UBound(Массив, 1)

Запрос = "" 'гост без года
If колич_элем <> 0 Then
  For i = 0 To колич_элем - 1
     Запрос = Запрос + Массив(i) + "-"
  Next i
Else
Запрос = ГОСТ_целеком
End If


'ИНТЕРНЕТ
If x_ГОСТ = СписокГОСТов.Count Or x_ГОСТ = 0 Then
  ie.Silent = True
  ie.Visible = True
  ie.Navigate "http://i1:8085/idoc/client/jsp/main.jsp?trail=~C~1~A~2~S~child_oks~C~3~A~2~S~child_oks~C~4~A~2~S~child_oks~C~403586~A~2#~C~1~A~2~S~child_oks~C~3~A~2~V~2~C~3~A~2"
  
  '
  Do While ie.Busy = True Or ie.ReadyState <> 4: DoEvents: Loop
End If

ie.Document.getElementById("findPatt").Value = Запрос

'Debug.Print ie.Document.getElementsByTagName("td")(2).getElementsByTagName("input")(0).getAttribute("alt")
'Debug.Print ie.Document.getElementsByTagName("input")(3).getAttribute("alt")

КолСимвДО = Len(ie.Document.body.innerText)


'Debug.Print КолСимвДО

'Ждем пока количество символов станен не 26155

'If x_ГОСТ = СписокГОСТов.Count Or x_ГОСТ = 0 Then
'  Продолжить = False
'  Do While Продолжить <> True
'   Ожидание = Len(ie.Document.body.innerText)
'   If Ожидание = 26155 Or Ожидание = 26336 Then Продолжить = True
'   DoEvents
'  Loop
'Else
  'Ждем некоторое время
  Время = Timer + 3 'число эта секунда
  Do While Timer < Время
   DoEvents
   'Debug.Print Len(ie.Document.body.innerText)
  Loop
'End If



'Ждем некоторое время
Время = Timer + 1 'число эта секунда
Do While Timer < Время
 DoEvents
 'Debug.Print Len(ie.Document.body.innerText)
Loop

КолСимвПосл = Len(ie.Document.body.innerText)
'Debug.Print КолСимвПосл

'КолСимвДО = Len(ie.Document.body.innerText)
'Debug.Print КолСимвДО



ie.Document.getElementsByTagName("input")(3).Click

'КолСимвВ = Len(ie.Document.body.innerText)
'Debug.Print КолСимвВ


'Ждем пока количество символов станен не 4005
Продолжить = False
Do While Продолжить <> True
 Ожидание = Len(ie.Document.body.innerText)
 If Ожидание <> 4005 Then Продолжить = True
 DoEvents
Loop

'Ждем некоторое время
Время = Timer + 2 'число эта секунда
Do While Timer < Время
 DoEvents
 'Debug.Print Len(ie.Document.body.innerText)
Loop

' КолСимвПосл = Len(ie.Document.body.innerText)
' Debug.Print КолСимвПосл
 

ТекстДокумента = ie.Document.body.innerText
If InStr(1, ТекстДокумента, "Ничего не найдено") <> 0 Then
  Результат_ГОСТ = ""
  Exit Sub
End If
'ТекстДокументаHTML = ie.Document.body.innerHTML
'Debug.Print ТекстДокумента
'Debug.Print ТекстДокументаHTML


'ВЫТАСКИВАЕМ САМЫЙ СВЕЖИЙ ГОД

Длина_запроса = Len(Запрос)
Свежий_Год = "0"

номер = 0
Год = "0"

Сново_Ищем_ГОД:
  
  If Год = "" Then Год = "0"
  If Свежий_Год = "" Then Свежий_Год = "0"
  If CSng(Replace(Свежий_Год, "-", "")) < CSng(Replace(Год, "-", "")) Then Свежий_Год = Год
  'On Error GoTo 0  'сново не пропускаем ошибки

  'Свежий_номер = номер
  Год = ""
  номер = 0
  Н_Перв_Симв = InStr(1 + Н_Перв_Симв, ТекстДокумента, Запрос)
  If Н_Перв_Симв = 0 Then GoTo Закончили_поиск_года
  Н_Последн_Симв = Н_Перв_Симв + Длина_запроса
  For i = Н_Последн_Симв To Len(ТекстДокумента)
    DoEvents
    Символ = Mid(ТекстДокумента, Н_Последн_Симв + номер, 1)
    If Символ Like "[0-9-]" Then 'любые числа и тирэ
       Год = Год + Символ
       номер = номер + 1
    Else 'если не цифра
      GoTo Сново_Ищем_ГОД
      'Exit For
    End If
    DoEvents
  Next i

Закончили_поиск_года:


Результат_ГОСТ = Запрос + Свежий_Год
If ПОЛНЫЙ_поиск = False Then
  Selection.Text = Результат
  Selection.Range.HighlightColorIndex = wdBrightGreen
End If


'ie.Document.getElementById("findTreePatt").Click
'Set g = ie.Document.getElementsByClassName("find")(1)
'ie.Document.getElementsByClassName("findTreePatt")(5).Click

'ie.Document.all("findPatt").Value = "3"
'ie.Visible = True
'ie.Navigate URL
''ie.Document.all.Item (10)
'ie.Document.doSimpleFind 1, Window.event.shiftKey
'ie.Document.all("findPatt").Click
'ie.Document.all("findPatt").focus

'Debug.Print ie.Document.getElementsByTagName("text")(1).getAttribute("alt")
'Debug.Print ie.Document.getElementById("pop_text").getAttribute("text")


'For Each o In ie.Document.getElementsByTagName("span")
'Debug.Print o.getAttribute("Text")
'Next

'ie.Document.getElementBytagname("findPatt").Value = "3"


'objshell.SendKeys "ГОСТ 2.729-68"
'objshell.SendKeys "{Enter}"
'objshell.SendKeys "~"
If ПОЛНЫЙ_поиск = False Then
  ie.Quit
  Set objshell = Nothing
  Set ie = Nothing
  Set СписокГОСТов = Nothing
End If


End Sub


Sub Проверка_годов_ГОСТов()
Dim r As Range, x, s$
t = Timer

Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=1, Name:="" 'перейти на стр по номеру
  
ReDim Массив_Брака(2) 'задаем размер таблицы
'Шаблоны битых гостов (беда с пробелами)
Массив_Брака(0) = "ГОСТ[0-9.-]{2;}"
Массив_Брака(1) = "<[О]СТ [0-9.-]{2;}"
Массив_Брака(2) = "<[Р]Д [0-9.-]{2;}"
'Ищем битые госты
For i = 0 To UBound(Массив_Брака, 1)
Dim cl As Collection
  Set r = ActiveDocument.Range
  Set cl = New Collection
  On Error Resume Next
  With r.Find
    .Text = Массив_Брака(i) 'ГОСТ [0-9.-]{2;} <[А-ЯЁ]{2;}>
    .Forward = True
    .Wrap = wdFindStop
    .Format = False
    .MatchWildcards = True
    While .Execute
      s = r.Text
      For x = 1 To cl.Count
        If s < cl(x) Then cl.Add s, s, Before:=x: GoTo 2
      Next
      cl.Add s, s
2   Wend
  End With

' Заменяем битые госты на правельные
For x_ГОСТ = cl.Count To 1 Step -1
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = cl.Item(x_ГОСТ)
        If i = 0 Then
          .Replacement.Text = Replace(cl.Item(x_ГОСТ), "ГОСТ", "ГОСТ ")
          .Replacement.Text = Replace(.Replacement.Text, "ГОСТ" & Chr(160), "ГОСТ ")
        End If
        If i = 1 Then
          .Replacement.Text = Replace(cl.Item(x_ГОСТ), "ОСТ ", "ОСТ")
          .Replacement.Text = Replace(.Replacement.Text, "ОСТ" & Chr(160), "ОСТ")
        End If
        If i = 2 Then
          .Replacement.Text = Replace(cl.Item(x_ГОСТ), "РД ", "РД")
          .Replacement.Text = Replace(.Replacement.Text, "РД" & Chr(160), "РД")
        End If
        
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
Next x_ГОСТ
Next i

Set cl = Nothing
  
  
'ИЩЕМ И ПРОВЕРЯЕМ ДАТУ У ГОСТОВ ПО ШАБЛНАМ
  номер = 9
  ReDim Массив_ГОСТов(номер) 'задаем размер таблицы
  номер = номер - номер
  'Шаблоны гостов
  Массив_ГОСТов(номер) = "ГОСТ [0-9.-]{2;}":         номер = номер + 1
  Массив_ГОСТов(номер) = "ГОСТ Р ИСО [0-9.-]{2;}":   номер = номер + 1
  Массив_ГОСТов(номер) = "ГОСТ [А-ЯЁ] [0-9.-]{2;}":  номер = номер + 1
  Массив_ГОСТов(номер) = "ГОСТ РВ [0-9.-]{2;}":      номер = номер + 1
  Массив_ГОСТов(номер) = "ОСТ[0-9.-]{2;}":           номер = номер + 1  'ГОСТ [0-9.Р]{2;}-[0-9]{2;}
  Массив_ГОСТов(номер) = "ОСТВ5Р[0-9.-]{2;}":        номер = номер + 1  'ГОСТ [0-9.Р]{2;}-[0-9]{2;}
  Массив_ГОСТов(номер) = "ОСТВ5[0-9.-]{2;}":         номер = номер + 1  'ГОСТ [0-9.Р]{2;}-[0-9]{2;}
  Массив_ГОСТов(номер) = "ОСТ5Р[0-9.-]{2;}":         номер = номер + 1  'ГОСТ [0-9.Р]{2;}-[0-9]{2;}
  Массив_ГОСТов(номер) = "ОСТ5[0-9.-]{2;}":          номер = номер + 1  'ГОСТ [0-9.Р]{2;}-[0-9]{2;}
  Массив_ГОСТов(номер) = "РД5[0-9.-]{2;}":           номер = номер + 1  'ГОСТ [0-9.Р]{2;}-[0-9]{2;}

  Set r = ActiveDocument.Range
  Set СписокГОСТов = New Collection
  On Error Resume Next
  
For i = 0 To UBound(Массив_ГОСТов, 1)
  With r.Find
    .Text = Массив_ГОСТов(i) 'ГОСТ [0-9.-]{2;} <[А-ЯЁ]{2;}>
    .Forward = True
    .Wrap = wdFindStop 'wdFindContinue 'wdFindStop
    If i > 0 Then
      .Wrap = wdFindContinue 'wdFindContinue 'wdFindStop
    Else
      .Wrap = wdFindStop 'wdFindContinue 'wdFindStop
    End If
    .Format = False
    .MatchWildcards = True
    While .Execute
      s = r.Text
      For x = 1 To СписокГОСТов.Count
        DoEvents
        If s < СписокГОСТов(x) Then СписокГОСТов.Add s, s, Before:=x: GoTo 1
      Next
      СписокГОСТов.Add s, s
1     .Wrap = wdFindStop 'wdFindContinue 'wdFindStop
   Wend
  End With
  
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
Next i

On Error GoTo 0  'сново не пропускаем ошибки


' ПОИСК в ГАЙКЕ


For x_ГОСТ = СписокГОСТов.Count To 1 Step -1
  
       Application.ScreenUpdating = True 'включить обновление документа
       DoEvents
       Application.ScreenUpdating = False 'включить обновление документа
       UserForm1.Label1.Width = CInt((300 * (СписокГОСТов.Count - x_ГОСТ + 1)) / СписокГОСТов.Count)
       UserForm1.Label_Лист = "Проверяемый документ: " & СписокГОСТов.Item(x_ГОСТ) & "  (" & СписокГОСТов.Count - x_ГОСТ + 1 & " из " & СписокГОСТов.Count & ")"
       UserForm1.Label2 = "Время: " & TimeSerial(0, 0, Timer - t)
       
       UserForm1.Repaint
  
  
  Искомый_ГОСТ = СписокГОСТов.Item(x_ГОСТ)
  ПОЛНЫЙ_поиск = True
  Call ГОСТ_Обновить_год
  ПОЛНЫЙ_поиск = False
  Что_Нашло = Результат_ГОСТ
  If Результат_ГОСТ = "" Then
    Options.DefaultHighlightColorIndex = wdRed
    Результат_ГОСТ = Искомый_ГОСТ
  Else
    Options.DefaultHighlightColorIndex = wdBrightGreen
  End If
  'СписокГОСТов.Item(x_ГОСТ) = Результат_ГОСТ
    
    Selection.Find.ClearFormatting
    Selection.Find.Highlight = False 'исходный текст не выделенцветом
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Highlight = True 'искомы текст выделить цветом
    'Selection.Find.Font.Color = 10498160
    With Selection.Find
        .Text = Искомый_ГОСТ
        .Replacement.Text = Результат_ГОСТ
        '.Replacement.Highlight = wdBrightGreen
        .Forward = True
        .Wrap = wdFindContinue
        'f = Selection.Find.Execute
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    
    Selection.Find.Execute Replace:=wdReplaceAll

Next
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting

End Sub



Sub ЗАГРУЗКА_Проверка_годов_ГОСТов()
  Application.ScreenUpdating = False 'выключить обновление документа
  'UserForm1.Repaint
   UserForm1.Label1.Width = 0
   UserForm1.Label_Лист = "Подготовка"
  'DoEvents
  ГАЛКА_Проверка_годов_ГОСТов = True
  UserForm1.Show
  ГАЛКА_Проверка_годов_ГОСТов = False
  'UserForm1.Hide
  On Error Resume Next
  ie.Quit
  Set objshell = Nothing
  Set ie = Nothing
  Set СписокГОСТов = Nothing
  On Error GoTo 0  'сново не пропускаем ошибки
  
  Application.ScreenUpdating = True 'включить обновление документа
End Sub






Sub Вставляем_лист_и_Заносим_в_колонтитул()
t = Timer
    
'If колонтитулы_листов_по_последн_изм = True Then
' GoTo колонтитулы_листов_по_последн
'End If
    
 ' If колонтитулы_всех_листов = False Then      'не работает макрос Берем_намера_стр_из_табл
 '    сообщение = MsgBox("Yes-""Зам."", No-""Нов."", Cancel-выйти ", vbYesNoCancel + vbQuestion + vbDefaultButton1, "Запрос на список страниц")
 '    If сообщение = vbCancel Then Exit Sub
 '    If сообщение = vbYes Then Лист = "Зам."
 '    If сообщение = vbNo Then Лист = "Нов."
 ' End If
  
  Лист = "Нов."
  
'колонтитулы_листов_по_последн:
    
 '   If колонтитулы_всех_листов = False Then
       Application.ScreenUpdating = False 'отключить обновление документа
 '   End If
 
 Колич_листов_в_документе = ActiveDocument.ComputeStatistics(wdStatisticPages)
    
 If Selection.PageSetup.PageWidth = CSng(Format(CentimetersToPoints(21), "0.0")) And _
      Selection.PageSetup.PageHeight = CSng(Format(CentimetersToPoints(29.7), "0.0")) And _
      Selection.PageSetup.VerticalAlignment = 0 _
 Then
      Формат_листа = "А4"
 Else
      Формат_листа = "А3"
 End If
    
    
If Формат_листа = "А4" Then
   текущий_масштаб = ActiveWindow.ActivePane.View.Zoom.Percentage
   ActiveWindow.ActivePane.View.Zoom.Percentage = 150
End If

   номер_активной_стр = Selection.Information(wdActiveEndPageNumber) 'номер активной стр

      
      Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=номер_активной_стр, Name:="" 'перейти на стр по номеру
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
      WordBasic.ViewFooterOnly 'перейти на нижний калантитул
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
      WordBasic.ViewFooterOnly 'перейти на нижний калантитул
      номер_листа_искомый_стар = Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text
      номер_листа_искомый_стар = Replace(номер_листа_искомый_стар, Chr(13) & "", "")
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
      
      Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=номер_активной_стр + 1, Name:="" 'перейти на стр по номеру
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
      WordBasic.ViewFooterOnly 'перейти на нижний калантитул
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
      WordBasic.ViewFooterOnly 'перейти на нижний калантитул
      номер_листа_впереди_стар = Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text
      номер_листа_впереди_стар = Replace(номер_листа_впереди_стар, Chr(13) & "", "")
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула

      If Формат_листа = "А4" Then
         номер_строки_старт = 3
      Else
         номер_строки_старт = 4
      End If

'если False заполняем толко текущую стр
 'If колонтитулы_всех_листов = False Then     'не работает макрос Берем_намера_стр_из_табл
   Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=номер_активной_стр, Name:="" 'перейти на стр по номеру
   With ActiveDocument.Range.Tables(ActiveDocument.Range.Tables.Count)
      Erase табл 'отчистить массив
      ReDim табл(.Rows.Count * 2, .Rows.Count * 2) 'задаем размер таблицы
      'последнее изменение
      For i = номер_строки_старт To .Rows.Count   ' получение стр из последней табл для каждого изм (а1, а2 ...)
         If .Cell(i, 1).Range.Text <> Chr(13) & "" Then
            табл(a1, 1) = Left(.Cell(i, 1).Range.Text, Len(.Cell(i, 1).Range.Text) - 2)
            'Лист = "Зам."
            табл(a1, 4) = Left(.Cell(i, 7).Range.Text, Len(.Cell(i, 7).Range.Text) - 2)
            табл(a1, 5) = Left(.Cell(i, 9).Range.Text, Len(.Cell(i, 9).Range.Text) - 2)
            табл(a1, 6) = Left(.Cell(i, 10).Range.Text, Len(.Cell(i, 10).Range.Text) - 2)
         End If
      Next i
   End With
' End If
 

'+++Вставляем разрыв страницы если его нет и снимаем галку "как в предыдущем разделе -  страница за текущей"

    
   ' Selection.MoveUp Unit:=wdLine, Count:=1, Extend:=wdExtend

    
  '  Selection.Find.ClearFormatting
  '  f = False
  '  With Selection.Find
  '      .Text = "^b"
  '      .Replacement.Text = ""
  '      .Forward = True
  '      .Wrap = wdFindStop 'wdFindStop не спарашивать если не нашло
  '       f = Selection.Find.Execute
  '      .Format = False
  '      .MatchCase = False
   '     .MatchWholeWord = False
  '      .MatchWildcards = False
  '      .MatchSoundsLike = False
  '      .MatchAllWordForms = False
 ''   End With
  '  If f = False Then
  '    Selection.MoveDown Unit:=wdLine, Count:=1
 '     Selection.MoveUp Unit:=wdLine, Count:=1
  '    Selection.InsertBreak Type:=wdSectionBreakContinuous  'разрыв на текущей странице на следующей стр
 '  Else
  '    Selection.MoveRight Unit:=wdCharacter, Count:=1
 '   End If
    
'+++Вставляем разрыв страницы если его нет и снимаем галку "как в предыдущем разделе -  страница перед текущей"
      'снимаем номер следуйщей стр
    Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=номер_активной_стр + 1, Name:="" 'перейти на стр по номеру
    Selection.MoveUp Unit:=wdLine, Count:=1, Extend:=wdExtend
    Selection.Find.ClearFormatting
    f = False
    With Selection.Find
        .Text = "^b"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop 'wdFindStop не спарашивать если не нашло
         f = Selection.Find.Execute
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    If f = False Then
      Selection.MoveDown Unit:=wdLine, Count:=1
      If Формат_листа = "А4" Then
         Selection.HomeKey Unit:=wdLine 'перейти в начало строки
         Selection.InsertBreak Type:=wdSectionBreakNextPage    'разрыв на следующей странице
      Else
         Selection.MoveUp Unit:=wdLine, Count:=1
         Selection.InsertBreak Type:=wdSectionBreakContinuous  'разрыв на текущей странице на следующей стр
      End If
    Else
     Selection.MoveRight Unit:=wdCharacter, Count:=1
    End If
      Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=номер_активной_стр + 1, Name:="" 'перейти на стр по номеру
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
      WordBasic.ViewFooterOnly 'перейти на нижний калантитул
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
      WordBasic.ViewFooterOnly 'перейти на нижний калантитул
      'WordBasic.ViewFooterOnly 'перейти на нижний калантитул
      Selection.HeaderFooter.LinkToPrevious = False ' снять "как в педыдущем разделе"
      номер_листа_впереди_нов = Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text
      номер_листа_впереди_нов = Replace(номер_листа_впереди_нов, "", "")
      номер_листа_впереди_нов = Replace(номер_листа_впереди_нов, Chr(13), "")
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула

      Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=номер_активной_стр, Name:="" 'перейти на стр по номеру
      WordBasic.ViewFooterOnly ' открыть нижний колонтитул
      
      номер_листа_искомый_нов = Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text
      номер_листа_искомый_нов = Replace(номер_листа_искомый_нов, "", "")
      номер_листа_искомый_нов = Replace(номер_листа_искомый_нов, Chr(13), "")
      номер_листа_искомый_нов_ВОРД = Selection.HeaderFooter.PageNumbers.StartingNumber
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
      
     'снимаем номер предыдущей стр
 '   If номер_активной_стр <> номер_активной_стр Then
  '    Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=номер_активной_стр - 1, Name:="" 'перейти на стр по номеру
  '    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
  '    WordBasic.ViewFooterOnly 'перейти на нижний калантитул
  '    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
  '    WordBasic.ViewFooterOnly 'перейти на нижний калантитул
  '    номер_листа_взади_нов = Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text
  '    номер_листа_взади_нов = Replace(номер_листа_взади_нов, "", "")
  '    номер_листа_взади_нов = Replace(номер_листа_взади_нов, Chr(13), "")
 '     ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
 '     Selection.GoTo What:=wdGoToPage, Which:=wdGoToNext, Name:="+1"
 '   End If
      
      ' номер_листа_взади_нов
      ' номер_листа_искомый_нов
      ' номер_листа_впереди_нов
      
      ' номер_листа_взади_стар
      ' номер_листа_искомый_стар
      ' номер_листа_впереди_стар

    '  If номер_листа_взади_нов <> номер_листа_взади_стар Then
    '        Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=номер_активной_стр - 1, Name:="" 'перейти на стр по номеру
    '       WordBasic.ViewFooterOnly ' открыть нижний колонтитул
    '         With Selection.HeaderFooter.PageNumbers  ' продолжить нумацию
    '           .NumberStyle = wdPageNumberStyleArabic
    '           .HeadingLevelForChapter = 0
    '           .IncludeChapterNumber = False
    '           .ChapterPageSeparator = wdSeparatorHyphen
    '           .RestartNumberingAtSection = False
    '           .StartingNumber = 0
   '         End With
   '         If Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text, Chr(13) & "", "") <> номер_листа_взади_стар Then
   '            With Selection.HeaderFooter.PageNumbers  ' задать номер стр номером
   '              .NumberStyle = wdPageNumberStyleArabic
   '              .HeadingLevelForChapter = 0
    '             .IncludeChapterNumber = False
   '              .ChapterPageSeparator = wdSeparatorHyphen
   '              .RestartNumberingAtSection = True
   '              .StartingNumber = номер_листа_взади_стар
   '           End With
   '        End If
   '        ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
   '   End If
      
      If номер_листа_искомый_нов <> номер_листа_искомый_стар Then
            Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=номер_активной_стр, Name:="" 'перейти на стр по номеру
           WordBasic.ViewFooterOnly ' открыть нижний колонтитул
             With Selection.HeaderFooter.PageNumbers  ' продолжить нумацию
               .NumberStyle = wdPageNumberStyleArabic
               .HeadingLevelForChapter = 0
               .IncludeChapterNumber = False
               .ChapterPageSeparator = wdSeparatorHyphen
               .RestartNumberingAtSection = False
               .StartingNumber = 0
            End With
            If Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text, Chr(13) & "", "") <> номер_листа_искомый_стар Then
               With Selection.HeaderFooter.PageNumbers  ' задать номер стр номером
                 .NumberStyle = wdPageNumberStyleArabic
                 .HeadingLevelForChapter = 0
                 .IncludeChapterNumber = False
                 .ChapterPageSeparator = wdSeparatorHyphen
                 .RestartNumberingAtSection = True
                 .StartingNumber = Список_страниц(номер_активной_стр - 1)
              End With
           End If
           If Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text, Chr(13) & "", "") <> номер_листа_искомый_стар Then
               Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text = номер_листа_искомый_стар
           End If
           ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
      End If
      
      
      'If номер_листа_впереди_нов <> номер_листа_впереди_стар Then
      ' фиксим значение стр на "номер_листа_впереди" стр
           Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=номер_активной_стр + 1, Name:="" 'перейти на стр по номеру
           WordBasic.ViewFooterOnly ' открыть нижний колонтитул
               With Selection.HeaderFooter.PageNumbers  ' задать номер стр номером
                 .NumberStyle = wdPageNumberStyleArabic
                 .HeadingLevelForChapter = 0
                 .IncludeChapterNumber = False
                 .ChapterPageSeparator = wdSeparatorHyphen
                 .RestartNumberingAtSection = True
                 .StartingNumber = номер_листа_впереди_стар
              End With
        '   End If
           ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
      'End If

      
'++++++++'ДОБЛЯЕМ НОВУЮ СТРАНИЦУ (набиваем строк пока они не убегут на следующую стр)

If Формат_листа = "А4" Then
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.EndKey Unit:=wdLine 'курсор в конец строки
    Selection.InsertBreak Type:=wdSectionBreakNextPage    'разрыв на следующей странице
    Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=номер_активной_стр + 1, Name:="" 'перейти на стр по номеру
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
    WordBasic.ViewFooterOnly 'перейти на нижний калантитул
    Selection.HeaderFooter.LinkToPrevious = False ' снять "как в педыдущем разделе"
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
    WordBasic.ViewFooterOnly 'перейти на нижний калантитул
Else
    'ActiveWindow.ActivePane.View.NextHeaderFooter 'следующий раздел
    'ActiveWindow.ActivePane.View.PreviousHeaderFooter 'предыдущий раздел
    
    'Определяем количество строк на странице
    номер_текущей_стр = номер_активной_стр
    колическо_строк_на_стр = 0
     Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=номер_активной_стр, Name:="" 'перейти на стр по номеру
     Do While номер_активной_стр + 1 <> номер_текущей_стр
          Selection.MoveDown Unit:=wdLine, Count:=1 'на абзац в вниз
          колическо_строк_на_стр = колическо_строк_на_стр + 1
          номер_текущей_стр = Selection.Information(wdActiveEndPageNumber) 'номер текущей стр
          DoEvents
     Loop
     колическо_строк_на_стр = колическо_строк_на_стр - 1
     

     
     'Добовляем строки в конце стартовой страницы
     Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=номер_активной_стр + 1, Name:="" 'перейти на стр по номеру
     Selection.MoveUp Unit:=wdLine, Count:=1 'на абзац в верх
     Selection.MoveUp Unit:=wdLine, Count:=1 'на абзац в верх
     
     'Колич_листов_в_докум = ActiveDocument.ComputeStatistics(wdStatisticPages)
     'If номер_активной_стр = Колич_листов_в_докум - 1 Then
     '   Selection.MoveUp Unit:=wdLine, Count:=1 'на абзац в верх
     'End If
     
    'Делаем рамку у последней строки табл на начальном листе
On Error Resume Next 'пропускаем ошибки
    Selection.Cells(1).Row.Select
    If Err.Number = 5991 Then 'Если ячейки в столбеце объеденены
       'определение номера активной таблицы
       Set tblSel = Selection.Tables(1)
       ingStart = tblSel.Range.Start
       For i = 1 To ActiveDocument.Tables.Count Step 1
          If ActiveDocument.Tables(i).Range.Start = ingStart Then
            ingTbIndex = i
            Exit For
          End If
       Next i
     'Выделяем свю строку вручныю
       For i = 1 To ActiveDocument.Tables(ingTbIndex).Columns.Count
         Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
       Next i
    End If
On Error GoTo 0  'сново не пропускаем ошибки
   ' With Selection.Borders(wdBorderTop)
   '     .LineStyle = Options.DefaultBorderLineStyle
   '     .LineWidth = Options.DefaultBorderLineWidth
   '     .Color = Options.DefaultBorderColor
   ' End With
   ' With Selection.Borders(wdBorderLeft)
   '     .LineStyle = Options.DefaultBorderLineStyle
   '     .LineWidth = Options.DefaultBorderLineWidth
   '     .Color = Options.DefaultBorderColor
   ' End With
    With Selection.Borders(wdBorderBottom)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
   ' With Selection.Borders(wdBorderRight)
   '     .LineStyle = Options.DefaultBorderLineStyle
   '     .LineWidth = Options.DefaultBorderLineWidth
   '     .Color = Options.DefaultBorderColor
   ' End With
   ' With Selection.Borders(wdBorderVertical)
   '     .LineStyle = Options.DefaultBorderLineStyle
   '     .LineWidth = Options.DefaultBorderLineWidth
    '    .Color = Options.DefaultBorderColor
   ' End With
     
   ' With Selection.Borders(wdBorderHorizontal)
   '     .LineStyle = Options.DefaultBorderLineStyle
   '     .LineWidth = Options.DefaultBorderLineWidth
   '     .Color = Options.DefaultBorderColor
   ' End With
    With Selection.Borders(wdBorderTop)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With

   'скрыть прапвую палку сетки таблицы
    Selection.Borders(wdBorderRight).LineStyle = wdLineStyleNone


     Selection.InsertRowsBelow колическо_строк_на_стр 'Вставить строки в таблицу
     'скрыть верхн палку сетки таблицы
     Selection.Borders(wdBorderTop).LineStyle = wdLineStyleNone
     
     'Проверка достаточно ли програ создала строк
     Do While номер_активной_стр + 1 = номер_текущей_стр
          Selection.InsertRowsBelow 1 'Вставить строки в таблицу
          'колическо_строк_на_стр = колическо_строк_на_стр + 1
          номер_текущей_стр = Selection.Information(wdActiveEndPageNumber) 'номер текущей стр
          DoEvents
     Loop
     Selection.Rows.Delete


    Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=номер_активной_стр + 1, Name:="" 'перейти на стр по номеру
    Selection.InsertBreak Type:=wdSectionBreakContinuous  'разрыв на текущей странице на следующей стр

     
   'снять "как в педыдущем разделе"
   WordBasic.ViewFooterOnly ' открыть нижний колонтитул
    If Selection.HeaderFooter.LinkToPrevious = True Then ' снять "как в педыдущем разделе"
       Selection.HeaderFooter.LinkToPrevious = False ' снять "как в педыдущем разделе"
       ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула если не выйти из колонтитула и сново не войти глючит и вписывает текст не втот колонтитул
       WordBasic.ViewFooterOnly ' открыть нижний колонтитул
       ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    End If

    'Делаем рамку у последней строки табл на начальном листе
     Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=номер_активной_стр + 1, Name:="" 'перейти на стр по номеру
     Selection.MoveUp Unit:=wdLine, Count:=1 'на абзац в верх
     Selection.MoveUp Unit:=wdLine, Count:=1 'на абзац в верх
     On Error Resume Next 'пропускаем ошибки
         Selection.Cells(1).Row.Select
         If Err.Number = 5991 Then 'Если ячейки в столбеце объеденены
            'определение номера активной таблицы
            Set tblSel = Selection.Tables(1)
            ingStart = tblSel.Range.Start
            For i = 1 To ActiveDocument.Tables.Count Step 1
               If ActiveDocument.Tables(i).Range.Start = ingStart Then
                 ingTbIndex = i
                 Exit For
               End If
            Next i
          'Выделяем свю строку вручныю
            For i = 1 To ActiveDocument.Tables(ingTbIndex).Columns.Count
              Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
            Next i
         End If
     On Error GoTo 0  'сново не пропускаем ошибки
     With Selection.Borders(wdBorderBottom)
         .LineStyle = Options.DefaultBorderLineStyle
         .LineWidth = Options.DefaultBorderLineWidth
         .Color = Options.DefaultBorderColor
     End With

End If 'Формат_листа = "А4"

    Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=номер_активной_стр + 1, Name:="" 'перейти на стр по номеру
    WordBasic.ViewFooterOnly ' открыть нижний колонтитул
    Selection.HeaderFooter.LinkToPrevious = False ' снять "как в педыдущем разделе"
  
'ReDim табл(1, 6) 'задаем размер таблицы
'табл(a1, 1) = "а1"
'Лист = "Зам."
'табл(a1, 4) = "22220.43.___"
'табл(a1, 5) = "Мазилевский"
'табл(a1, 6) = "15.16.18"


'Ищем точке исходной номере стр
дельта_после_запятой = 0.1
перед_запятой = номер_листа_искомый_нов
If InStr(1, номер_листа_искомый_нов, ".") > 0 Then
   после_запятой = CSng(Split(номер_листа_искомый_нов, ".")(1))
   перед_запятой = CSng(Split(номер_листа_искомый_нов, ".")(0))
   дельта_после_запятой = 10 ^ (-1 * CSng(Len(после_запятой)))
End If
'Меняем номер стр
'Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Select
'Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:="PAGE  \* Arabic ", PreserveFormatting:=True

'Вписываем номер страницы
  Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text = CStr(перед_запятой) + "."
  Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Select
  Selection.EndKey Unit:=wdLine 'курсор в конец строки
  Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:="PAGE  \* Arabic ", PreserveFormatting:=True  'вордовский умный номер страницы



If InStr(1, номер_листа_искомый_нов, ".") = 0 Then  'если на стартовой стр небыло числа после точки в номере стр
               With Selection.HeaderFooter.PageNumbers  ' задать номер стр номером
                 .NumberStyle = wdPageNumberStyleArabic
                 .HeadingLevelForChapter = 0
                 .IncludeChapterNumber = False
                 .ChapterPageSeparator = wdSeparatorHyphen
                 .RestartNumberingAtSection = True
                 .StartingNumber = 1
              End With
            
            If Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text, Chr(13) & "", "") = Replace(номер_листа_искомый_нов + дельта_после_запятой, ",", ".") Then
               GoTo Конец_Цикла
            End If

Else 'если было
             With Selection.HeaderFooter.PageNumbers  ' продолжить нумацию
               .NumberStyle = wdPageNumberStyleArabic
               .HeadingLevelForChapter = 0
               .IncludeChapterNumber = False
               .ChapterPageSeparator = wdSeparatorHyphen
               .RestartNumberingAtSection = False
               .StartingNumber = 0
            End With
            
            If Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text, Chr(13) & "", "") = Replace(номер_листа_искомый_нов + дельта_после_запятой, ",", ".") Then
               GoTo Конец_Цикла
            End If
            
               With Selection.HeaderFooter.PageNumbers  ' задать номер стр номером
                 .NumberStyle = wdPageNumberStyleArabic
                 .HeadingLevelForChapter = 0
                 .IncludeChapterNumber = False
                 .ChapterPageSeparator = wdSeparatorHyphen
                 .RestartNumberingAtSection = True
                 .StartingNumber = после_запятой + 1
              End With
              
            If Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text, Chr(13) & "", "") = Replace(номер_листа_искомый_нов + дельта_после_запятой, ",", ".") Then
               GoTo Конец_Цикла
            End If
            
End If

Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text = Replace(номер_листа_искомый_нов + дельта_после_запятой, ",", ".")

Конец_Цикла:

'If Replace(после_запятой + 1, ",", ".") = номер_листа_искомый_нов_ВОРД Then
'  Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Text = перед_запятой + "."
'  Selection.EndKey Unit:=wdLine 'курсор в конец строки
'  Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:="PAGE  \* Arabic ", PreserveFormatting:=True
  'Selection.HomeKey Unit:=wdLine 'курсор в начало строки
  'Selection.TypeText Text:=перед_запятой
'Else
'  Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text = Replace(номер_листа_искомый_нов + дельта_после_запятой, ",", ".")
'End If
'номер_листа_искомый_нов






'Selection.HeaderFooter.PageNumbers.StartingNumber 'номер стр назначеный ворде (Номерация страниц/формат номеров страниц....)



'заполняем колонтитул текстом
  If подтереть_колон = False Then
   '   Selection.Find.Execute 'выделить начденый текст  'Dim табл()
      'Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 1).Range.Text = табл(a1, 1) 'таблица в колонтитуле
      'Selection.Tables(1).Cell(2, 1).Range.Text = "а1" 'таблица в колонтитуле
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 2).Range.Text = Лист 'таблица в колонтитуле
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 3).Range.Text = табл(a1, 4) 'таблица в колонтитуле
      'Selection.Tables(1).Cell(2, 3).Range.Text = "22220.43.___" 'таблица в колонтитуле
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 3).Select
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 3).FitText = True
      'Selection.HeaderFooter.Range.Cells(1).FitText = True  'сузить до ширены ячейки
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 4).Range.Text = табл(a1, 5) 'таблица в колонтитуле
      'Selection.Tables(1).Cell(2, 4).Range.Text = "Мазилевский" 'таблица в колонтитуле
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 4).Select
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 4).FitText = True  'сузить до ширены ячейки
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 5).Range.Text = табл(a1, 6) 'таблица в колонтитуле
      'Selection.Tables(1).Cell(2, 5).Range.Text = "15.16.18" 'таблица в колонтитуле
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 5).Select
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 5).FitText = True
  End If
  

      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' выйти из колонтитула
  
    
    If (Timer - t) > 60 Then
      Debug.Print (Timer - t) / 60 & " мин" ' время  в сек
      Else
      Debug.Print Timer - t & " сек" ' время  в сек
    End If
  
    
к_выходу:

If Формат_листа = "А4" Then
    ActiveWindow.ActivePane.View.Zoom.Percentage = текущий_масштаб
End If

Application.ScreenUpdating = True 'включить обновление документа
    

End Sub



Sub Разрыв_табл()
Attribute Разрыв_табл.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.тест"


'определение номера активной таблицы
Set tblSel = Selection.Tables(1)
ingStart = tblSel.Range.Start
For i = 1 To ActiveDocument.Tables.Count Step 1
 If ActiveDocument.Tables(i).Range.Start = ingStart Then
    табл_ingTbIndex = i
    Exit For
 End If
Next i
 'определение координат октивной ячейки
 табл_у1 = Selection.Information(wdEndOfRangeRowNumber) 'номер строки ячейки таблицы к которой тыкнут курсор
 табл_х1 = Selection.Information(wdEndOfRangeColumnNumber) 'номер столбца ячейки таблицы к которой тыкнут курсор
   
   Call Разрыв_табл_клавиша

End Sub

Sub Разрыв_табл_клавиша()
    'Назначаем горячие главиши
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyReturn), _
        KeyCategory:=wdKeyCategoryCommand, Command:="Разрыв_табл_Вставляем_шапку"
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyEsc), _
        KeyCategory:=wdKeyCategoryCommand, Command:="Разрыв_табл_Вставляем_шапку_Отмена"
        
        MsgBox ("Выделите шапку таблицы и нажмите Enter." & vbNewLine & vbNewLine & "Для отмены нажмите ОК затем Esc")
        
End Sub

Sub Разрыв_табл_Вставляем_шапку_Отмена()
'Снять горячию клавишу
CustomizationContext = NormalTemplate
FindKey(BuildKeyCode(Arg1:=wdKeyReturn)).Clear
CustomizationContext = NormalTemplate
FindKey(BuildKeyCode(Arg1:=wdKeyEsc)).Clear
Err.Clear
End Sub


Sub Разрыв_табл_Вставляем_шапку()
Attribute Разрыв_табл_Вставляем_шапку.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.тест"

On Error Resume Next 'пропускаем ошибки
    Selection.Copy
      ActiveDocument.Tables(табл_ingTbIndex).Cell(табл_у1 - 1, табл_х1).Select ' фокусируемся на одну строчку ниже выбраной
      Selection.HomeKey Unit:=wdLine 'перейти в начало строки (ячейки)
      'Selection.SplitTable ' разбить таблицу
      'Selection.InsertBreak Type:=wdSectionBreakNextPage  'разрыв на следующей странице на следующей стр
      Selection.InsertBreak Type:=wdPageBreak 'обычный разрыв страницы
      ActiveDocument.Tables(табл_ingTbIndex + 1).Cell(1, 1).Select
      Selection.HomeKey Unit:=wdLine 'перейти в начало строки
    Selection.Paste
    ActiveDocument.Tables(табл_ingTbIndex + 1).Cell(1, 1).Select
    Selection.InsertRowsAbove 1 'Добавить строчку наверх
    Selection.Cells.Merge ' Объединить ячейку
    ' Рамка
    Selection.Borders(wdBorderTop).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderLeft).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderRight).LineStyle = wdLineStyleNone
    
    Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft 'выравнивание по левому краю
    Selection.TypeText Text:="Продолжение таблицы " 'вписать текст
    
'Снять горячию клавишу
CustomizationContext = NormalTemplate
FindKey(BuildKeyCode(Arg1:=wdKeyReturn)).Clear
CustomizationContext = NormalTemplate
FindKey(BuildKeyCode(Arg1:=wdKeyEsc)).Clear
Err.Clear

On Error GoTo 0  'сново не пропускаем ошибки
End Sub

Sub Далой_Экспоненту()
    'Чистем формат искомого текста
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Highlight = False
    Selection.Find.Replacement.Text = ""
    
    'Ищем степень в экспоненциальном виде и выделяем
    Options.DefaultHighlightColorIndex = wdPink ' цвет менять здесь
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindContinue
        .Text = "E+[0-9]{2;}"
        .Replacement.Highlight = True
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=Word.wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindContinue
        .Text = "E-[0-9]{2;}"
        .Replacement.Highlight = True
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=Word.wdReplaceAll
    End With
    
   'удаляем лишний ноль после плюса или минуса
    Selection.Find.ClearFormatting
    Selection.Find.Highlight = True
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "E+0"
        .Replacement.Text = "E"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Highlight = True
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "E-0"
        .Replacement.Text = "E-"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
   'делаем раскрашеные числа надстрочными и снимаем раскраску
    Selection.Find.ClearFormatting
    Selection.Find.Highlight = True
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Highlight = False
    With Selection.Find.Replacement.Font
        .Superscript = True
        .Subscript = False
    End With
    With Selection.Find
        .Text = "[0-9,-]{1;}"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    'заменяем раскрашеные E+0 на ·10
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Highlight = True
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Highlight = False
    With Selection.Find
        .Text = "E+"
        .Replacement.Text = "·10"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=Word.wdReplaceAll
    End With
    
    'заменяем раскрашеные E на 10
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Highlight = True
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Highlight = False
    With Selection.Find
        .Text = "E"
        .Replacement.Text = "·10"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=Word.wdReplaceAll
    End With
    
    'Чистем формат искомого текста
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Highlight = False
    Selection.Find.Replacement.Text = ""
    
End Sub



Sub AutoExec()
'Sub AutoExec()
    '''''востановить сочетания клавиш по умолчанию
    '''''CustomizationContext = NormalTemplate
    '''''KeyBindings.ClearAll
'''If Application.CommandBars.Count > 203 Then Exit Sub

'Удаляю все кнопки
On Error Resume Next 'пропускаем ошибки
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
Application.CommandBars("Меню").Delete
On Error GoTo 0  'сново не пропускаем ошибки

  With Application.CommandBars.Add("Меню", tamporary = True)
     With .Controls.Add
     .Caption = "Добавить в колонтитул последнее изм."
     .FaceId = 2063
     .Style = 3
     .TooltipText = "Добавить в колонтитул последнее изм. из таблицы листа регистрации изменений"
     .OnAction = "Заносим_в_колонтитул"
     End With
     .Visible = True
  End With
  
'Вставить лист и занести изв. в колонтит.
  With Application.CommandBars.Add("Меню", tamporary = True)
     With .Controls.Add
     .Caption = "Вставить лист и занести изв. в колонтит."
     .FaceId = 9419 '3145
     .Style = 3
     .TooltipText = ""
     .OnAction = "Вставляем_лист_и_Заносим_в_колонтитул"
     End With
     .Visible = True
  End With
  
  With Application.CommandBars.Add("Меню", tamporary = True)
     With .Controls.Add
     .Caption = "Вносим в колонтитулы ВСЕ изм."
     .FaceId = 303
     .Style = 3
     .TooltipText = "Вносим в колонтитулы все изм. по ""Листу регистрации изменений"" (первый, последний не корректируются; листы без таблиц выдадут ошибку, поэтому предварительно удалите глюченые листы из списков). Не понимает нецелые номера листов через тире (типа:3.1-3.3)."
     .OnAction = "ЗАГРУЗКА_Проверка_и_замена_всех_колонт"
     End With
     .Visible = True
  End With
  
  With Application.CommandBars.Add("Меню", tamporary = True)
     With .Controls.Add
     .Caption = "Сделать/не сделать фиолетовым, жирным текст"
     .FaceId = 1382
     .Style = 3
     .TooltipText = "Ctrl+W Сделать/не сделать фиолетовым, жирным выделенный текст"
     .OnAction = "Фиолетовый"
     End With
     .Visible = True
  End With
  
  With Application.CommandBars.Add("Меню", tamporary = True)
     With .Controls.Add
     .Caption = "Растянуть текст по ширине ячейки табл."
     .FaceId = 1355 '542
     .Style = 3
     .TooltipText = "Ctrl+E Растягиваем текст в выделенной ячейке по ширине"
     .OnAction = "ПоШирене"
     End With
     .Visible = True
  End With
  
  With Application.CommandBars.Add("Меню", tamporary = True)
     With .Controls.Add
     .Caption = "Сохранить в файл TXT полный список листов"
     .FaceId = 1548 '139
     .Style = 3
     On Error Resume Next 'пропускаем ошибки
     .TooltipText = "Файл TXT сохраняется в той же папке, что и файл документа под именем """ & ActiveDocument.Name & "_страницы.txt"""
     On Error GoTo 0  'сново не пропускаем ошибки
     .OnAction = "ЗАГРУЗКА_Список_стр_в_файл"
     End With
     .Visible = True
  End With
  
  With Application.CommandBars.Add("Меню", tamporary = True)
     With .Controls.Add
     .Caption = "Список страниц c несоответствующем колонтитулом"
     .FaceId = 1446
     .Style = 3
     .TooltipText = "Показать список страниц у которых колонтитулы отличаются от указанных в ""Листе регистрации изменений"""
     .OnAction = "ЗАГРУЗКА_Сравнение_колонтитулов_в_листах_с_указаными_в_таблице"
     End With
     .Visible = True
  End With
  
  With Application.CommandBars.Add("Меню", tamporary = True)
     With .Controls.Add
     .Caption = "Удалить большенство разрывов и колонтитулов"
     .FaceId = 3202 '214'1716
     .Style = 3
     .TooltipText = "Удалить большенство разрывов и колонтитулов (листы с дробной нумерацией не трогает). Список страниц в TXT должен быть актуальным."
     .OnAction = "ЗАГРУЗКА_Удаляем_разрывы"
     End With
     .Visible = True
  End With
  
  With Application.CommandBars.Add("Меню", tamporary = True)
     With .Controls.Add
     .Caption = "Справка"
     .FaceId = 984 '214'1716
     .Style = 3
     .TooltipText = ""
     .OnAction = "Инструкция"
     End With
     .Visible = True
  End With
  
'Красим_в_Зеленый
  With Application.CommandBars.Add("Меню", tamporary = True)
     With .Controls.Add
     .Caption = "Красим в Зеленым"
     .FaceId = 6735 '394
     .Style = 3
     .TooltipText = "Ctrl+1 Выделеный текст делаем в зеленом выделении"
     .OnAction = "Красим_в_Зеленый"
     End With
     .Visible = True
  End With
  
'Красим_в_Желтый_1
  With Application.CommandBars.Add("Меню", tamporary = True)
     With .Controls.Add
     .Caption = "Красим в Желтый"
     .FaceId = 6751 '351
     .Style = 3
     .TooltipText = "Ctrl+2 Выделеный текст делаем в желтом выделении"
     .OnAction = "Красим_в_Желтый_1"
     End With
     .Visible = True
  End With
  
'Красим_в_Красный
  With Application.CommandBars.Add("Меню", tamporary = True)
     With .Controls.Add
     .Caption = "Красим в Красный"
     .FaceId = 6743 '352
     .Style = 3
     .TooltipText = "Ctrl+3 Выделеный текст делаем в красном выделении"
     .OnAction = "Красим_в_Красный"
     End With
     .Visible = True
  End With
  
'Заменяем тоячуки на запятые
  With Application.CommandBars.Add("Меню", tamporary = True)
     With .Controls.Add
     .Caption = "Замена точек на запятые"
     .FaceId = 382 '
     .Style = 3
     .TooltipText = "Заменяет точки и запятые, удаляет пробелы в таблице спецификации с столбцах: код, количество и общ. сумма"
     .OnAction = "ЗАГРУЗКА_Замена_точек_на_запятые"
     End With
     .Visible = True
  End With
  
'Вносим последнее изм в в колонтитуле по всему докум по последней табл (список регистр измен)
  With Application.CommandBars.Add("Меню", tamporary = True)
     With .Controls.Add
     .Caption = "Вносим в колонтитулы последнее изм."
     .FaceId = 159 '
     .Style = 3
     .TooltipText = "Вносим последнее изм из табл регистр. измен. на все листы указаные в этой табл"
     .OnAction = "ЗАГРУЗКА_Вносим_в_колонтитулы_только_последние_изменения"
     End With
     .Visible = True
  End With
  
'Удалить из ленты надстройку
  With Application.CommandBars.Add("Меню", tamporary = True)
     With .Controls.Add
     .Caption = "Удалить из ленты надстройку"
     .FaceId = 4305 '330 1088
     .Style = 3
     .TooltipText = "Удалить из ленты надстройку"
     .OnAction = "Удалить_ленту"
     End With
     .Visible = True
  End With

'Пересчетать общ. колич. во всех пунках
  With Application.CommandBars.Add("Меню", tamporary = True)
     With .Controls.Add
     .Caption = "Пересчетать общ. массу во всех пунках"
     .FaceId = 283 '
     .Style = 3
     .TooltipText = "Пересчетать общую массу во всех пунках документа"
     .OnAction = "ЗАГРУЗКА_Пересчет_всех_пунктов"
     End With
     .Visible = True
  End With
  
'Канкулятор в ячейке
  With Application.CommandBars.Add("Меню", tamporary = True)
     With .Controls.Add
     .Caption = "Канкулятор в ячейке табл."
     .FaceId = 50 '
     .Style = 3
     .TooltipText = "Ctrl+R Конкулятор в ячейке табл."
     .OnAction = "Канкулятор_в_ячейке"
     End With
     .Visible = True
  End With
  
'Рассчитать общую массу в строке
  With Application.CommandBars.Add("Меню", tamporary = True)
     With .Controls.Add
     .Caption = "Рассчитать общую массу в строке"
     .FaceId = 385
     .Style = 3
     .TooltipText = "Ctrl+Q Перед тем как нажать на кнопку - тыкните курсором в строку где надо получить общую массу позиции (произведение количества на массу единицы)"""
     .OnAction = "Сумм_Масса"
     End With
     .Visible = True
  End With
  
'Список сокращений
  With Application.CommandBars.Add("Меню", tamporary = True)
     With .Controls.Add
     .Caption = "Список сокращений"
     .FaceId = 1031 '
     .Style = 3
     .TooltipText = "Список сокращений появится в конце текста"
     .OnAction = "А_Сокращения"
     End With
     .Visible = True
  End With
  
'Проверить ГОСТы
  With Application.CommandBars.Add("Меню", tamporary = True)
     With .Controls.Add
     .Caption = "Проверить ГОСТы"
     .FaceId = 1922 '202
     .Style = 3
     .TooltipText = "ГОСТам пишет правельный год"
     .OnAction = "ЗАГРУЗКА_Проверка_годов_ГОСТов"
     End With
     .Visible = True
  End With
  
'Разрыв табл
  With Application.CommandBars.Add("Меню", tamporary = True)
     With .Controls.Add
     .Caption = "Разорвать таблицу"
     .FaceId = 2233 '635
     .Style = 3
     .TooltipText = "Тыкаем мышкой на табл где ее нужно разорвать"
     .OnAction = "Разрыв_табл"
     End With
     .Visible = True
  End With
  
'Далой Экспоненту
  With Application.CommandBars.Add("Меню", tamporary = True)
     With .Controls.Add
     .Caption = "Далой Экспоненту"
     .FaceId = 57 '
     .Style = 3
     .TooltipText = "Заменяет Е+3 на 10^3"
     .OnAction = "Далой_Экспоненту"
     End With
     .Visible = True
  End With
  
  Application.CommandBars.Add("Меню", tamporary = True).Controls.Add
    
    'Назначаем горячие главиши
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyQ, wdKeyControl), _
        KeyCategory:=wdKeyCategoryCommand, Command:="Сумм_Масса"
        
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyE, wdKeyControl), _
        KeyCategory:=wdKeyCategoryCommand, Command:="ПоШирене"
        
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyW, wdKeyControl), _
        KeyCategory:=wdKeyCategoryCommand, Command:="Фиолетовый"
    'Красим_в_Зеленый
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKey1, wdKeyControl), _
        KeyCategory:=wdKeyCategoryCommand, Command:="Красим_в_Зеленый"
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyNumeric1, wdKeyControl), _
        KeyCategory:=wdKeyCategoryCommand, Command:="Красим_в_Зеленый"
    'Красим_в_Желтый_1
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKey2, wdKeyControl), _
        KeyCategory:=wdKeyCategoryCommand, Command:="Красим_в_Желтый_1"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyNumeric2, wdKeyControl), _
        KeyCategory:=wdKeyCategoryCommand, Command:="Красим_в_Желтый_1"
    'Красим_в_Красный
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKey3, wdKeyControl), _
        KeyCategory:=wdKeyCategoryCommand, Command:="Красим_в_Красный"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyNumeric3, wdKeyControl), _
        KeyCategory:=wdKeyCategoryCommand, Command:="Красим_в_Красный"
    'Красим_в_Голубой
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKey4, wdKeyControl), _
        KeyCategory:=wdKeyCategoryCommand, Command:="Красим_в_Голубой"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyNumeric4, wdKeyControl), _
        KeyCategory:=wdKeyCategoryCommand, Command:="Красим_в_Голубой"
    'Красим_в_Серый
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKey5, wdKeyControl), _
        KeyCategory:=wdKeyCategoryCommand, Command:="Красим_в_Серый"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyNumeric5, wdKeyControl), _
        KeyCategory:=wdKeyCategoryCommand, Command:="Красим_в_Серый"
    'Канкулятор в ячейке
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyR, wdKeyControl), _
        KeyCategory:=wdKeyCategoryCommand, Command:="Канкулятор_в_ячейке"
    
End Sub


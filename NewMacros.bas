Attribute VB_Name = "NewMacros"
'����� ����������� ���� ��������
'��������� ��� ������ ���������������� � ������� �3

Option Base 0

Dim ����()
Dim a1
Dim ���� As String
Dim �����������_����_������, �����_��������_���
Public ������_�������(), ���������_�����, ������_���(), ������_�����_���(), ����_�������
Public �����_��������_�_������_����_������, �����_������_���_�_����, �����������_������_��_�������_��� As Boolean
Public �����_���������_������������_�_������_�_���������_�_�������, �����_�������_������� As Boolean
Public �����_������_�����_��_�������, �����_������_�_�����������_������_���������_���������, �����_��������_����_�������, ������_�����, �����_��������_�����_������ As Boolean
Public �������_����, ���������_����, x_����
Public ������������ As Collection
Public ����_�1, ����_�1, ����_ingTbIndex
Public ie As InternetExplorerMedium



Sub ��������()
Attribute ��������.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.��������"
'
' �������� ������
'
'
On Error Resume Next '���������� ������
    'Selection.Cells(1).Select
 ������ = Selection.Cells(1).FitText
If ������ = True Then
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
    Selection.HomeKey Unit:=wdLine '������� � ������ ������
Else
    Selection.Cells(1).FitText = True
End If
End Sub


Sub ����_�����()
'
' ����_����� ������
'
'
On Error Resume Next '���������� ��� ������


�1 = Selection.Rows.First.Index  '����� �������� ������ � �������

'����������� ������ �������� �������
Set tblSel = Selection.Tables(1)
ingStart = tblSel.Range.Start
For i = 1 To ActiveDocument.Tables.Count Step 1
 If ActiveDocument.Tables(i).Range.Start = ingStart Then
    ingTbIndex = i
    Exit For
 End If
 Next i
 


���������� = ActiveDocument.Tables(ingTbIndex).Cell(�1, 5).Range.Text    '������� �������� � ������ � �������
���������� = Replace(����������, Chr(13) & "", "")
���������� = Replace(����������, " ", "")
���������� = Replace(����������, vbrTab, "")
���������� = Replace(����������, Chr(9), "")
���������� = Replace(����������, ".", ",")

'���������� = Left(����������, Len(����������) - 1)                       '������� ��������� ���� - ������ ��� ������
������� = ActiveDocument.Tables(ingTbIndex).Cell(�1, 6).Range            '������� �������� � ������ � �������
������� = Replace(�������, Chr(13) & "", "")
������� = Replace(�������, " ", "")
������� = Replace(�������, vbrTab, "")
������� = Replace(�������, Chr(9), "")
'������� = Left(�������, Len(�������) - 1)                                '������� ��������� ���� - ������ ��� ������

'If InStr(1, CStr(����������), ".") <> 0 Or InStr(1, CStr(����������), ",") <> 0 Then
'   ����������_������_�����_��� = ����������
'End If

'���������� ���������� ������ ����� �������
����������_������_�����_��� = 0
If InStr(1, CStr(����������), ".") <> 0 Or InStr(1, CStr(����������), ",") <> 0 Then
   If InStr(1, CStr(����������), ".") <> 0 Then
      ����������_������_�����_��� = Len(Split(����������, ".")(1))
   End If
   If InStr(1, CStr(����������), ",") <> 0 Then
      ����������_������_�����_��� = Len(Split(����������, ",")(1))
   End If
End If

If InStr(1, CStr(�������), ".") <> 0 Or InStr(1, CStr(�������), ",") <> 0 Then
   If InStr(1, CStr(�������), ".") <> 0 Then
      If ����������_������_�����_��� < Len(Split(�������, ".")(1)) Then
         ����������_������_�����_��� = Len(Split(�������, ".")(1))
      End If
   End If
   If InStr(1, CStr(�������), ",") <> 0 Then
      If ����������_������_�����_��� < Len(Split(�������, ",")(1)) Then
         ����������_������_�����_��� = Len(Split(�������, ",")(1))
      End If
   End If
End If

'���������� = Replace(����������, ".", Application.International(xlDecimalSeparator)) ' ������ ����� �� �������
'������� = Replace(�������, ".", Application.International(xlDecimalSeparator))    ' ������ ����� �� �������

���������� = Replace(����������, ",", ".")  ' ������ ����� �� �������
������� = Replace(�������, ",", ".")    ' ������ ����� �� �������

If ���������� = "" Or ���������� = "0" Then Exit Sub
If ������� = "" Or ������� = "0" Then Exit Sub


��������� = ���������� * �������

If ��������� = 0 Then
  ���������� = Replace(����������, ".", ",")  ' ������ ����� �� �������
  ������� = Replace(�������, ".", ",")    ' ������ ����� �� �������
  ��������� = ���������� * �������
End If
If ��������� = 0 Then ��������� = ""


'���������
'If ��������� >= 0 And ��������� < 0.00001 Then ��������� = Format(���������, "0.0000000")
'If ��������� >= 0.00001 And ��������� < 0.0001 Then ��������� = Format(���������, "0.000000")
'If ��������� >= 0.0001 And ��������� < 0.001 Then ��������� = Format(���������, "0.00000")
'If ��������� >= 0.001 And ��������� < 0.01 Then ��������� = Format(���������, "0.00000")
'If ��������� >= 0.01 And ��������� < 1 Then ��������� = Format(���������, "0.0000")
'If ��������� >= 1 And ��������� < 10 Then ��������� = Format(���������, "0.00")
'If ��������� >= 10 And ��������� < 500 Then ��������� = Format(���������, "0.0")
'If ��������� >= 500 Then ��������� = Format(���������, "0.0")
'��������� = Round(���������, ����������_������_�����_���)

'���������� ������ ����� �������
If ����������_������_�����_��� = 0 Then ��������� = Format(���������, "0")
If ����������_������_�����_��� = 1 Then ��������� = Format(���������, "0.0")
If ����������_������_�����_��� = 2 Then ��������� = Format(���������, "0.00")
If ����������_������_�����_��� = 3 Then ��������� = Format(���������, "0.000")
If ����������_������_�����_��� = 4 Then ��������� = Format(���������, "0.0000")
If ����������_������_�����_��� = 5 Then ��������� = Format(���������, "0.00000")
If ����������_������_�����_��� = 6 Then ��������� = Format(���������, "0.000000")
If ����������_������_�����_��� = 7 Then ��������� = Format(���������, "0.0000000")

' ������� � ������
��������� = Replace(���������, ".", ",")  ' ������ �������  �� �����
ActiveDocument.Tables(ingTbIndex).Cell(�1, 7).Range = ���������

'���������� ���������� ������ ����� ������� "���������"
����������_������_�����_���_��������� = 0
If InStr(1, CStr(���������), ".") <> 0 Or InStr(1, CStr(���������), ",") <> 0 Then
   If InStr(1, CStr(���������), ".") <> 0 Then
      ����������_������_�����_���_��������� = Len(Split(���������, ".")(1))
   End If
   If InStr(1, CStr(���������), ",") <> 0 Then
      ����������_������_�����_���_��������� = Len(Split(���������, ",")(1))
   End If

'���������� ���������� ������ ����� ������� "���������"
����������_������_�����_���_��������� = 0
   If InStr(1, CStr(���������), ".") <> 0 Then
      ����������_������_�����_���_��������� = Len(Split(���������, ".")(0))
   End If
   If InStr(1, CStr(���������), ",") <> 0 Then
      ����������_������_�����_���_��������� = Len(Split(���������, ",")(0))
   End If
Else
    ����������_������_�����_���_��������� = Len(���������)
End If


If ����������_������_�����_���_��������� + ����������_������_�����_���_��������� > 7 Then
ActiveDocument.Tables(ingTbIndex).Cell(�1, 7).Range.Select
Call ��������
End If

End Sub
Sub ����������()
Attribute ����������.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.����������"
'
' ���������� ������
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
 
' ��������� � �������� ��������� ������ ���� ",,5,6,8,,9-15,18,2,11-9,,1,4,,21,"
    ' ���������� ���������� (��������������) ������ � �������
    ' array(5,6,8,9,10,11,12,13,14,15,18,2,11,10,9,1,4,21)
    ' (������ �������� ���������; ��������� ���� 9-15 � 17-13 ������������)

arr = Split(Replace(Txt$, " ", ""), ","): Dim n As Long: ReDim tmpArr(0 To 0)
    For i = LBound(arr) To UBound(arr)
        Select Case True
            Case arr(i) = "", Val(arr(i)) < 0
                '  ���������������� ��� ������, ����� ������ � ������� ��������
                '  ���� ����������� � ��������� (����������������� � �������� -1)
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
 

Sub �������_�_����������()
t = Timer
    
If �����������_������_��_�������_��� = True Then
 GoTo �����������_������_��_�������
End If
    
  If �����������_����_������ = False Then      '�� �������� ������ �����_������_���_��_����
     ��������� = MsgBox("Yes-""���."", No-""���."", Cancel-����� ", vbYesNoCancel + vbQuestion + vbDefaultButton1, "������ �� ������ �������")
     If ��������� = vbCancel Then Exit Sub
     If ��������� = vbYes Then ���� = "���."
     If ��������� = vbNo Then ���� = "���."
  End If
  
�����������_������_��_�������:
    
 If �����������_����_������ = False Then
       Application.ScreenUpdating = False '��������� ���������� ���������
 End If

 If Selection.PageSetup.PageWidth = CSng(Format(CentimetersToPoints(21), "0.0")) And _
      Selection.PageSetup.PageHeight = CSng(Format(CentimetersToPoints(29.7), "0.0")) And _
      Selection.PageSetup.VerticalAlignment = 0 _
 Then
      ������_����� = "�4"
 Else
      ������_����� = "�3"
 End If

    
If ������_����� = "�4" Then
   �������_������� = ActiveWindow.ActivePane.View.Zoom.Percentage
   ActiveWindow.ActivePane.View.Zoom.Percentage = 150
End If
    

    'Application.ScreenUpdating = False '��������� ���������� ���������
' ��������� ������ � ������, ���������� ��� �������� �������� ������
    
'+++++������� ������ ������� ������
'���� = ActiveDocument.Path
'��� = ActiveDocument.Name
'Open ���� & "\" & ��� & "_��������.txt" For Input As #1  '������� ��� ������ "Input"
'Line Input #1, �����_��_txt '������ ������ #1
'Close #1 '������� ��������

'������ = Split(�����_��_txt, ",") '�������� ������
'ReDim ������_�������(UBound(������, 1)) '������ ������ �������
'����������� �������� ����� ������� �������

'For i = 0 To UBound(������, 1)
'   ������_�������(i) = ������(i)
'Next i
   '����� �������� ��� �� � �����������
If �����������_������_��_�������_��� = False Then '����� �� ��������� � ���������� �� ����������� ����� �������� ���
   �����_��������_��� = Selection.Information(wdActiveEndPageNumber) '����� �������� ���
End If

      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
      WordBasic.ViewFooterOnly '������� �� ������ �� ��������
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
      WordBasic.ViewFooterOnly '������� �� ������ ����������
'If �����������_������_��_�������_��� = False Then

'End If
      ' �����_�����_�����_����
      ' �����_�����_�������_����
      ' �����_�����_�������_����
' If �����������_����_������ = False Then     '�� �������� ������ �����_������_���_��_����
      Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=�����_��������_��� - 1, Name:="" '������� �� ��� �� ������
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
      WordBasic.ViewFooterOnly '������� �� ������ �� ��������
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
      WordBasic.ViewFooterOnly '������� �� ������ ����������
      If Selection.Information(wdActiveEndPageNumber) = 1 Then
         �����_�����_�����_���� = "1"
      Else
         �����_�����_�����_���� = Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text
         �����_�����_�����_���� = Replace(�����_�����_�����_����, Chr(13) & "", "")
      End If
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
      
      Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=�����_��������_���, Name:="" '������� �� ��� �� ������
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
      WordBasic.ViewFooterOnly '������� �� ������ ����������
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
      WordBasic.ViewFooterOnly '������� �� ������ ����������
      �����_�����_�������_���� = Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text
      �����_�����_�������_���� = Replace(�����_�����_�������_����, Chr(13) & "", "")
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
      
      Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=�����_��������_��� + 1, Name:="" '������� �� ��� �� ������
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
      WordBasic.ViewFooterOnly '������� �� ������ ����������
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
      WordBasic.ViewFooterOnly '������� �� ������ ����������
      �����_�����_�������_���� = Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text
      �����_�����_�������_���� = Replace(�����_�����_�������_����, Chr(13) & "", "")
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
'End If

' If �����������_����_������ = True Then     '�������� ������ �����_������_���_��_����
'       �����_�����_�����_���� = ������_�������(�����_��������_��� - 2)
'       �����_�����_�������_���� = ������_�������(�����_��������_��� - 1)
'       �����_�����_�������_���� = ������_�������(�����_��������_���)
' End If

'���� False ��������� ����� ������� ���
 If �����������_����_������ = False Then     '�� �������� ������ �����_������_���_��_����
   Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=�����_��������_���, Name:="" '������� �� ��� �� ������
   With ActiveDocument.Range.Tables(ActiveDocument.Range.Tables.Count)
      Erase ���� '��������� ������
      ReDim ����(.Rows.Count * 2, .Rows.Count * 2) '������ ������ �������
      
      If ������_����� = "�4" Then
         �����_������_����� = 3
      Else
         �����_������_����� = 4
      End If
      '��������� ���������
      For i = �����_������_����� To .Rows.Count  ' ��������� ��� �� ��������� ���� ��� ������� ��� (�1, �2 ...)
         If .Cell(i, 1).Range.Text <> Chr(13) & "" Then
            ����(a1, 1) = Left(.Cell(i, 1).Range.Text, Len(.Cell(i, 1).Range.Text) - 2)
            '���� = "���."
            ����(a1, 4) = Left(.Cell(i, 7).Range.Text, Len(.Cell(i, 7).Range.Text) - 2)
            ����(a1, 5) = Left(.Cell(i, 9).Range.Text, Len(.Cell(i, 9).Range.Text) - 2)
            ����(a1, 6) = Left(.Cell(i, 10).Range.Text, Len(.Cell(i, 10).Range.Text) - 2)
         End If
      Next i
   End With
 End If
 

'+++��������� ������ �������� ���� ��� ��� � ������� ����� "��� � ���������� ������� -  �������� �� �������"

    
    Selection.MoveUp Unit:=wdLine, Count:=1, Extend:=wdExtend

    
    Selection.Find.ClearFormatting
    f = False
    With Selection.Find
        .Text = "^b"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop 'wdFindStop �� ����������� ���� �� �����
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
      If ������_����� = "�4" Then
         Selection.HomeKey Unit:=wdLine '������� � ������ ������
         Selection.InsertBreak Type:=wdSectionBreakNextPage    '������ �� ��������� ��������
      Else
         Selection.MoveUp Unit:=wdLine, Count:=1
         Selection.InsertBreak Type:=wdSectionBreakContinuous  '������ �� ������� �������� �� ��������� ���
      End If
    Else
      Selection.MoveRight Unit:=wdCharacter, Count:=1
    End If
    
'+++��������� ������ �������� ���� ��� ��� � ������� ����� "��� � ���������� ������� -  �������� ����� �������"
      '������� ����� ��������� ���
    Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=�����_��������_��� + 1, Name:="" '������� �� ��� �� ������
    Selection.MoveUp Unit:=wdLine, Count:=1, Extend:=wdExtend
    Selection.Find.ClearFormatting
    f = False
    With Selection.Find
        .Text = "^b"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop 'wdFindStop �� ����������� ���� �� �����
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
      If ������_����� = "�4" Then
         Selection.HomeKey Unit:=wdLine '������� � ������ ������
         Selection.InsertBreak Type:=wdSectionBreakNextPage    '������ �� ��������� ��������
      Else
         Selection.MoveUp Unit:=wdLine, Count:=1
         Selection.InsertBreak Type:=wdSectionBreakContinuous  '������ �� ������� �������� �� ��������� ���
      End If
    Else
     Selection.MoveRight Unit:=wdCharacter, Count:=1
    End If
      Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=�����_��������_��� + 1, Name:="" '������� �� ��� �� ������
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
      WordBasic.ViewFooterOnly '������� �� ������ ����������
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
      WordBasic.ViewFooterOnly '������� �� ������ ����������
      'WordBasic.ViewFooterOnly '������� �� ������ ����������
      Selection.HeaderFooter.LinkToPrevious = False ' ����� "��� � ��������� �������"
      �����_�����_�������_��� = Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text
      �����_�����_�������_��� = Replace(�����_�����_�������_���, "", "")
      �����_�����_�������_��� = Replace(�����_�����_�������_���, Chr(13), "")
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������

      Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=�����_��������_���, Name:="" '������� �� ��� �� ������
      WordBasic.ViewFooterOnly ' ������� ������ ����������
      
      �����_�����_�������_��� = Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text
      �����_�����_�������_��� = Replace(�����_�����_�������_���, "", "")
      �����_�����_�������_��� = Replace(�����_�����_�������_���, Chr(13), "")
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
      
     '������� ����� ���������� ���
    If �����_��������_��� <> �����_��������_��� Then
      Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=�����_��������_��� - 1, Name:="" '������� �� ��� �� ������
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
      WordBasic.ViewFooterOnly '������� �� ������ ����������
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
      WordBasic.ViewFooterOnly '������� �� ������ ����������
      �����_�����_�����_��� = Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text
      �����_�����_�����_��� = Replace(�����_�����_�����_���, "", "")
      �����_�����_�����_��� = Replace(�����_�����_�����_���, Chr(13), "")
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
      Selection.GoTo What:=wdGoToPage, Which:=wdGoToNext, Name:="+1"
    End If
      
      ' �����_�����_�����_���
      ' �����_�����_�������_���
      ' �����_�����_�������_���
      
      ' �����_�����_�����_����
      ' �����_�����_�������_����
      ' �����_�����_�������_����

      If �����_�����_�����_��� <> �����_�����_�����_���� And �����_��������_��� <> �����_��������_��� Then
            Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=�����_��������_��� - 1, Name:="" '������� �� ��� �� ������
           WordBasic.ViewFooterOnly ' ������� ������ ����������
             With Selection.HeaderFooter.PageNumbers  ' ���������� �������
               .NumberStyle = wdPageNumberStyleArabic
               .HeadingLevelForChapter = 0
               .IncludeChapterNumber = False
               .ChapterPageSeparator = wdSeparatorHyphen
               .RestartNumberingAtSection = False
               .StartingNumber = 0
            End With
            If Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text, Chr(13) & "", "") <> �����_�����_�����_���� Then
               With Selection.HeaderFooter.PageNumbers  ' ������ ����� ��� �������
                 .NumberStyle = wdPageNumberStyleArabic
                 .HeadingLevelForChapter = 0
                 .IncludeChapterNumber = False
                 .ChapterPageSeparator = wdSeparatorHyphen
                 .RestartNumberingAtSection = True
                 .StartingNumber = �����_�����_�����_����
              End With
           End If
           ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
      End If
      
      If �����_�����_�������_��� <> �����_�����_�������_���� Then
            Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=�����_��������_���, Name:="" '������� �� ��� �� ������
           WordBasic.ViewFooterOnly ' ������� ������ ����������
             With Selection.HeaderFooter.PageNumbers  ' ���������� �������
               .NumberStyle = wdPageNumberStyleArabic
               .HeadingLevelForChapter = 0
               .IncludeChapterNumber = False
               .ChapterPageSeparator = wdSeparatorHyphen
               .RestartNumberingAtSection = False
               .StartingNumber = 0
            End With
            If Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text, Chr(13) & "", "") <> �����_�����_�������_���� Then
               With Selection.HeaderFooter.PageNumbers  ' ������ ����� ��� �������
                 .NumberStyle = wdPageNumberStyleArabic
                 .HeadingLevelForChapter = 0
                 .IncludeChapterNumber = False
                 .ChapterPageSeparator = wdSeparatorHyphen
                 .RestartNumberingAtSection = True
                 .StartingNumber = ������_�������(�����_��������_��� - 1)
              End With
           End If
           If Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text, Chr(13) & "", "") <> �����_�����_�������_���� Then
               Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text = �����_�����_�������_����
           End If
           ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
      End If
      
      
      If �����_�����_�������_��� <> �����_�����_�������_���� Then
            Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=�����_��������_��� + 1, Name:="" '������� �� ��� �� ������
           WordBasic.ViewFooterOnly ' ������� ������ ����������
             With Selection.HeaderFooter.PageNumbers  ' ���������� �������
               .NumberStyle = wdPageNumberStyleArabic
               .HeadingLevelForChapter = 0
               .IncludeChapterNumber = False
               .ChapterPageSeparator = wdSeparatorHyphen
               .RestartNumberingAtSection = False
               .StartingNumber = 0
            End With
            If Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text, Chr(13) & "", "") <> �����_�����_�������_���� Then
               With Selection.HeaderFooter.PageNumbers  ' ������ ����� ��� �������
                 .NumberStyle = wdPageNumberStyleArabic
                 .HeadingLevelForChapter = 0
                 .IncludeChapterNumber = False
                 .ChapterPageSeparator = wdSeparatorHyphen
                 .RestartNumberingAtSection = True
                 .StartingNumber = �����_�����_�������_����
              End With
           End If
           ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
      End If

      
'++++++++������� ��������

    'ActiveWindow.ActivePane.View.NextHeaderFooter '��������� ������
    'ActiveWindow.ActivePane.View.PreviousHeaderFooter '���������� ������
     Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=�����_��������_���, Name:="" '������� �� ��� �� ������
     
   WordBasic.ViewFooterOnly ' ������� ������ ����������
    If Selection.HeaderFooter.LinkToPrevious = True Then ' ����� "��� � ��������� �������"
       Selection.HeaderFooter.LinkToPrevious = False ' ����� "��� � ��������� �������"
       ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� ����������� ���� �� ����� �� ����������� � ����� �� ����� ������ � ��������� ����� �� ���� ����������
       WordBasic.ViewFooterOnly ' ������� ������ ����������
    End If

  
'ReDim ����(1, 6) '������ ������ �������
'����(a1, 1) = "�1"
'���� = "���."
'����(a1, 4) = "22220.43.___"
'����(a1, 5) = "�����������"
'����(a1, 6) = "15.16.18"

'��������� ���������� �������
  If ���������_����� = False Then
   '   Selection.Find.Execute '�������� �������� �����  'Dim ����()
      'Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 1).Range.Text = ����(a1, 1) '������� � �����������
      'Selection.Tables(1).Cell(2, 1).Range.Text = "�1" '������� � �����������
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 2).Range.Text = ���� '������� � �����������
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 3).Range.Text = ����(a1, 4) '������� � �����������
      'Selection.Tables(1).Cell(2, 3).Range.Text = "22220.43.___" '������� � �����������
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 3).Select
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 3).FitText = True
      'Selection.HeaderFooter.Range.Cells(1).FitText = True  '������ �� ������ ������
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 4).Range.Text = ����(a1, 5) '������� � �����������
      'Selection.Tables(1).Cell(2, 4).Range.Text = "�����������" '������� � �����������
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 4).Select
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 4).FitText = True  '������ �� ������ ������
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 5).Range.Text = ����(a1, 6) '������� � �����������
      'Selection.Tables(1).Cell(2, 5).Range.Text = "15.16.18" '������� � �����������
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 5).Select
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 5).FitText = True
  End If
  
'������ ���������� ������
  If ���������_����� = True Then
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 1).Range.Text = "" '������� � �����������
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 2).Range.Text = "" '������� � �����������
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 3).Range.Text = "" '������� � �����������
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 4).Range.Text = "" '������� � �����������
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 5).Range.Text = "" '������� � �����������
  End If

 
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
  
    
    
    If (Timer - t) > 60 Then
      Debug.Print (Timer - t) / 60 & " ���" ' �����  � ���
      Else
      Debug.Print Timer - t & " ���" ' �����  � ���
    End If
    
    If �����������_����_������ = True Or �����������_������_��_�������_��� = True Then
       GoTo �_������
    End If
    
If ������_����� = "�4" Then
    ActiveWindow.ActivePane.View.Zoom.Percentage = �������_�������
End If

    
    Application.ScreenUpdating = True '�������� ���������� ���������
    
�_������:
    
    'Application.ScreenUpdating = True '�������� ���������� ���������
End Sub


Sub ������_���_�_����()
'Dim ������_�������()
  t = Timer
  
  Application.ScreenUpdating = False '��������� ���������� ���������
  
���� = ActiveDocument.Path
��� = ActiveDocument.Name
���_����� = ��� & "_��������.txt"



'��������� = MsgBox("��������/������� ������ ������� � �����:" & Chr(13) & """" & ���_����� & """", vbYesNoCancel + vbQuestion + vbDefaultButton1, "������ �� ������ �������")

If ��������� = vbCancel Then
  Application.ScreenUpdating = True
  Exit Sub
End If


  
  '����� ������� � ����������
For i = 1 To ActiveDocument.Tables.Count Step 1
   If ActiveDocument.Tables(i).Range.Columns.Count = 18 Then
      �����_������� = i
      Exit For
   End If
Next i

ActiveDocument.Tables(�����_�������).Range.Cells(1).Select
�����_���_�_������_���� = Selection.Information(wdActiveEndPageNumber) '����� �������� ���
 
'If ��������� = vbYes Then
  Erase ������_������� '��������� ������
  ReDim ������_�������(�����_���_�_������_����) '������ ������ �������
  
  For i = 0 To �����_���_�_������_���� - 2
     ������_�������(i) = Str(i + 1)
  Next i
  
  �����_������_�_����� = ActiveDocument.ComputeStatistics(wdStatisticPages)
  Selection.HomeKey Unit:=wdStory '������� � ������ ���������
  For j = �����_���_�_������_���� - 1 To ActiveDocument.ComputeStatistics(wdStatisticPages) - 1 '���������� �������
       
       �����_��������_��� = Selection.Information(wdActiveEndPageNumber) '����� �������� ���
       DoEvents
       UserForm1.Label1.Width = CInt((300 * �����_��������_���) / �����_������_�_�����)
       UserForm1.Label_���� = "���� " & �����_��������_��� & " �� " & �����_������_�_�����
       'If (Timer - t) > 60 Then
       '   UserForm1.Label2 = "����� " & Round((Timer - t) / 60, 1) & " ���" ' �����  � ���
       'Else
       '   UserForm1.Label2 = "����� " & Round(Timer - t, 1) & " ���" ' �����  � ���
       'End If
       UserForm1.Label2 = "�����: " & TimeSerial(0, 0, Timer - t)
       
       UserForm1.Repaint
              
      Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=j + 1, Name:="" '������� �� ��� �� ������
      WordBasic.ViewFooterOnly ' ������� ������ ����������
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
      WordBasic.ViewFooterOnly ' ������� ������ ����������
      ReDim Preserve ������_�������(j)
      ������_�������(j) = Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text '��������� � ������ � ���� � ������ �����������
      ������_�������(j) = Replace(������_�������(j), "", "")
      ������_�������(j) = Replace(������_�������(j), Chr(13), "")
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
  Next j
'End If 'If ��������� = vbYes Then


'�������� ���������� �� ����
'If Dir(���� & "\" & ���_�����) = "" Then
'������� ���� (���� ���� ��� ���� �� ������������ �� ������)
'    Set fso = CreateObject("scripting.filesystemobject")
'    Set ts = fso.createtextfile(���� & "\" & ��� & "_��������.txt", True)
'    ts.write txt: ts.Close
'    Set ts = Nothing: Set fso = Nothing
'End If

'������� � ������� ���������� ���� � ���������
    Open ���� & "\" & ���_����� For Output As #1
    Print #1, Join(������_�������, ",")
    Close #1


    Application.ScreenUpdating = True '�������� ���������� ���������
    'End If
    If (Timer - t) > 60 Then
      Debug.Print (Timer - t) / 60 & " ���" ' �����  � ���
      Else
      Debug.Print Timer - t & " ���" ' �����  � ���
    End If
End Sub




Sub �������_�����()

    '''''����������� ��������� ������ �� ���������
    CustomizationContext = NormalTemplate
    KeyBindings.ClearAll

'������ ��� ������
On Error Resume Next '���������� ������
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
On Error GoTo 0  '����� �� ���������� ������
End Sub





Sub ��������_������_���_�����������()
Dim ������_���_���(), ������_���_���(), ������_���_���
Dim ���_���, ���_���, ���_��� '�����



  
  �����_��������_��� = Selection.Information(wdActiveEndPageNumber) '����� �������� ���
  If �����_��������_��� = 1 Then Exit Sub
  
  '�������  ����� �������� ��������
  WordBasic.ViewFooterOnly ' ������� ������ ����������
  �����_��������_���_��_������ = Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text, Chr(13) & "", "")
  ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
  
  �����_������_�_����� = ActiveDocument.ComputeStatistics(wdStatisticPages)
  
  With ActiveDocument.Range.Tables(ActiveDocument.Range.Tables.Count)
    Erase ������_��� '��������� ������
    'ReDim ������_���(.Rows.Count * 2, �����_������_�_�����, �����_������_�_�����, �����_������_�_�����, 1, 1, 1) '������ ������ �������
    ReDim ������_���(�����_������_�_����� * 3, 5) '������ ������ �������  ������ � 3 ���� ������ ����� ����� ��������� ���� ��� ��������� ������ �����
    '������_�������(0) = "1"
 '������ ���������
 �����_��� = 0
 For i = 4 To .Rows.Count  ' ��������� ��� �� ��������� ���� ��� ������� ��� (�1, �2 ...)
    '.Cell(i, 1).Select
    If .Cell(i, 1).Range.Text <> Chr(13) & "" Then
        �����_��� = �����_��� + 1
    End If
 Next i
 
    �����_����� = 0
    ���_��� = ""
    ���_��� = ""
    ���_��� = ""
    g = 1
      For i = 4 To .Rows.Count  ' ��������� ��� �� ��������� ���� ��� ������� ��� (�1, �2 ...)
         '.Cell(i, 1).Select
        
         If .Cell(i, 1).Range.Text <> Chr(13) & "" Or .Cell(i, 2).Range.Text <> Chr(13) & "" _
         Or .Cell(i, 3).Range.Text <> Chr(13) & "" Or .Cell(i, 4).Range.Text <> Chr(13) & "" Then
           '.Cell(i, 1).Select
           
           If .Cell(i, 1).Range.Text <> Chr(13) & "" Then
               �����_����� = �����_����� + 1
           End If
           
           
           If .Cell(i, 1).Range.Text <> Chr(13) & "" Then
               �����_��� = Left(.Cell(i, 1).Range.Text, Len(.Cell(i, 1).Range.Text) - 2)
           End If
               ���_��� = ���_��� + Left(.Cell(i, 2).Range.Text, Len(.Cell(i, 2).Range.Text) - 2)
               ���_��� = ���_��� + Left(.Cell(i, 3).Range.Text, Len(.Cell(i, 3).Range.Text) - 2)
               ���_��� = ���_��� + Left(.Cell(i, 4).Range.Text, Len(.Cell(i, 4).Range.Text) - 2)
               �����_����� = �����_����� + Left(.Cell(i, 7).Range.Text, Len(.Cell(i, 7).Range.Text) - 2)
               ������� = ������� + Left(.Cell(i, 9).Range.Text, Len(.Cell(i, 9).Range.Text) - 2)
               ���� = ���� + Left(.Cell(i, 10).Range.Text, Len(.Cell(i, 10).Range.Text) - 2)
         End If
         
         
         If (.Cell(i + 1, 1).Range.Text <> Chr(13) & "" And �����_����� <> 0) Or (i = .Rows.Count) Then  '��� Or (i = .Rows.Count) ������� ��� ����� ���������� ���
               '��������� �������� ������ ��� � ������� (��������� �������� ���������� ������ ����������)
               ������_���_��� = ArrayOfValues(���_���)
               ������_���_��� = ArrayOfValues(���_���)
               ������_���_��� = ArrayOfValues(���_���)
               '������ �������� ����� �����
               If UBound(������_���_���, 1) >= UBound(������_���_���, 1) And UBound(������_���_���, 1) >= UBound(������_���_���, 1) Then
                 ReDim Preserve ������_���_���(UBound(������_���_���, 1))
                 ReDim Preserve ������_���_���(UBound(������_���_���, 1))
               End If
               If UBound(������_���_���, 1) >= UBound(������_���_���, 1) And UBound(������_���_���, 1) >= UBound(������_���_���, 1) Then
                 ReDim Preserve ������_���_���(UBound(������_���_���, 1))
                 ReDim Preserve ������_���_���(UBound(������_���_���, 1))
               End If
               If UBound(������_���_���, 1) >= UBound(������_���_���, 1) And UBound(������_���_���, 1) >= UBound(������_���_���, 1) Then
                 ReDim Preserve ������_���_���(UBound(������_���_���, 1))
                 ReDim Preserve ������_���_���(UBound(������_���_���, 1))
               End If
               
               For Y = 0 To (UBound(������_���_���, 1))
                  On Error Resume Next
                  ������_���_���(Y) = CSng(������_���_���(Y)) '���������  ����� � ������� (������� �������)
                  ������_���_���(Y) = CStr(������_���_���(Y)) '���������  � �����
                  ������_���_���(Y) = Replace(������_���_���(Y), ",", ".")
                  ������_���_���(Y) = CSng(������_���_���(Y)) '���������  ����� � ������� (������� �������)
                  ������_���_���(Y) = CStr(������_���_���(Y)) '���������  � �����
                  ������_���_���(Y) = Replace(������_���_���(Y), ",", ".")
                  ������_���_���(Y) = CSng(������_���_���(Y)) '���������  ����� � ������� (������� �������)
                  ������_���_���(Y) = CStr(������_���_���(Y)) '���������  � �����
                  ������_���_���(Y) = Replace(������_���_���(Y), ",", ".")
                  On Error GoTo 0
               Next Y
               '������� ������ ��������� ��� ������ �������� � ����������� � ������
               '�����_��� = 0
               '������_���_���
               For Y = 0 To (UBound(������_���_���, 1))
                  If ������_���_���(Y) <> 0 Then
                     ����� = False
                     '���������� ������ � ������� �������� � ����� �������
                     For h = 0 To (UBound(������_���, 1))
                        If ������_���(h, 0) = ������_���_���(Y) Then
                           ������_���(h, 0) = ������_���_���(Y)
                           ������_���(h, 1) = �����_���
                           ������_���(h, 2) = ""
                           ������_���(h, 3) = �����_�����
                           ������_���(h, 4) = �������
                           ������_���(h, 5) = ����
                           '�����_��� = �����_��� + 1
                           ����� = True '����� ����������
                           Exit For '����� � ������� ����� �� ����� �������
                        End If
                     Next h
                        If ����� = False Then '�� ����� ����������
                           ������_���(�����_���, 0) = ������_���_���(Y)
                           ������_���(�����_���, 1) = �����_���
                           ������_���(�����_���, 2) = ""
                           ������_���(�����_���, 3) = �����_�����
                           ������_���(�����_���, 4) = �������
                           ������_���(�����_���, 5) = ����
                           �����_��� = �����_��� + 1
                        End If
                  End If
               Next Y
               '������_���_���
               For Y = 0 To (UBound(������_���_���, 1))
                  If ������_���_���(Y) <> 0 Then
                     ����� = False
                     '���������� ������ � ������� �������� � ����� �������
                     For h = 0 To (UBound(������_���, 1))
                        If ������_���(h, 0) = ������_���_���(Y) Then
                           ������_���(h, 0) = ������_���_���(Y)
                           ������_���(h, 1) = �����_���
                           ������_���(h, 2) = "���."
                           ������_���(h, 3) = �����_�����
                           ������_���(h, 4) = �������
                           ������_���(h, 5) = ����
                           '�����_��� = �����_��� + 1
                           ����� = True '����� ����������
                           Exit For
                        End If
                     Next h
                        If ����� = False Then '�� ����� ����������
                           ������_���(�����_���, 0) = ������_���_���(Y)
                           ������_���(�����_���, 1) = �����_���
                           ������_���(�����_���, 2) = "���."
                           ������_���(�����_���, 3) = �����_�����
                           ������_���(�����_���, 4) = �������
                           ������_���(�����_���, 5) = ����
                           �����_��� = �����_��� + 1
                        End If
                  End If
               Next Y
               '������_���_���
               For Y = 0 To (UBound(������_���_���, 1))
                  If ������_���_���(Y) <> 0 Then
                     ����� = False
                     '���������� ������ � ������� �������� � ����� �������
                     For h = 0 To (UBound(������_���, 1))
                        If ������_���(h, 0) = ������_���_���(Y) Then
                           ������_���(h, 0) = ������_���_���(Y)
                           ������_���(h, 1) = �����_���
                           ������_���(h, 2) = "���."
                           ������_���(h, 3) = �����_�����
                           ������_���(h, 4) = �������
                           ������_���(h, 5) = ����
                           '�����_��� = �����_��� + 1
                           ����� = True '����� ����������
                           Exit For
                        End If
                     Next h
                        If ����� = False Then '�� ����� ����������
                           ������_���(�����_���, 0) = ������_���_���(Y)
                           ������_���(�����_���, 1) = �����_���
                           ������_���(�����_���, 2) = "���."
                           ������_���(�����_���, 3) = �����_�����
                           ������_���(�����_���, 4) = �������
                           ������_���(�����_���, 5) = ����
                           �����_��� = �����_��� + 1
                        End If
                  End If
               Next Y
           '��������
           ���_��� = ""
           ���_��� = ""
           ���_��� = ""
           �����_����� = ""
           ������� = ""
           ���� = ""
           End If
         
      Next i
   End With
  

 
End Sub

Sub ��������_�_������_����_������()
  ����� = Timer
  
  DoEvents
  
'����� ������� � ����������
For i = 1 To ActiveDocument.Tables.Count Step 1
   If ActiveDocument.Tables(i).Range.Columns.Count = 18 Then
      �����_������� = i
      Exit For
   End If
Next i

ActiveDocument.Tables(�����_�������).Range.Cells(1).Select
�����_���_�_������_���� = Selection.Information(wdActiveEndPageNumber) '����� �������� ���
   
 
  'UserForm1.Label1.Width = CInt((300 * �����_��������_���) / �����_������_�_�����)
  'UserForm1.Label_���� = "���� " & �����_��������_��� & " �� " & �����_������_�_�����
  UserForm1.Label1.Width = 0
  UserForm1.Repaint
  
 
  
  Application.ScreenUpdating = False '��������� ���������� ���������
  'UserForm1.Label1.Width = 0
  'UserForm1.Show
    �����_������_�_����� = ActiveDocument.ComputeStatistics(wdStatisticPages)
    With ActiveDocument.Range.Tables(ActiveDocument.Range.Tables.Count)
       Erase ������_��� '��������� ������
      'ReDim ������_���(.Rows.Count * 2, �����_������_�_�����, �����_������_�_�����, �����_������_�_�����, 1, 1, 1) '������ ������ �������
       ReDim ������_���(�����_������_�_����� * 3, 5) '������ ������ �������
    End With
    
Call ��������_������_���_�����������  '������_���()
  
  With ActiveDocument.Range.Tables(ActiveDocument.Range.Tables.Count)
      Erase ���� '��������� ������
      ReDim ����(.Rows.Count * 2, .Rows.Count * 2) '������ ������ �������
  End With
  a1 = 0
  

  For i = 2 To �����_������_�_�����

       '���������� ����� � �����������
       If i < �����_���_�_������_���� Then i = �����_���_�_������_����
       
       Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=i, Name:="" '������� �� ��� �� ������
       WordBasic.ViewFooterOnly ' ������� ������ ����������
       ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
       WordBasic.ViewFooterOnly ' ������� ������ ����������
       �����_��������_���_��_������ = Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text, Chr(13) & "", "")
       ���_��������_���_��_������ = Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 1).Range.Text, Chr(13) & "", "")
       ����_��������_���_��_������ = Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 2).Range.Text, Chr(13) & "", "")
       �����_�����_��������_���_��_������ = Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 3).Range.Text, Chr(13) & "", "")
       
       ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
       
       �����_��������_��� = Selection.Information(wdActiveEndPageNumber) '����� �������� ���
       'Application.ScreenUpdating = True
       'Application.StatusBar = "���� " & �����_��������_��� & " �� " & �����_������_�_�����
       'Application.StatusBar = False
       'Application.ScreenUpdating = False
       'Application.ScreenUpdating = True '�������� ���������� ���������
       DoEvents
       UserForm1.Label1.Width = CInt((300 * �����_��������_���) / �����_������_�_�����)
       UserForm1.Label_���� = "���� " & �����_��������_��� & " �� " & �����_������_�_�����
       'If (Timer - �����) > 60 Then
       '   UserForm1.Label2 = "����� " & Round((Timer - �����) / 60, 1) & " ���" ' �����  � ���
       'Else
       '   UserForm1.Label2 = "����� " & Round(Timer - �����, 1) & " ���" ' �����  � ���
       'End If
       UserForm1.Label2 = "�����: " & TimeSerial(0, 0, Timer - �����)
       
       UserForm1.Repaint
       'Application.ScreenUpdating = False '�������� ���������� ���������
       
       �����_��� = False
       For Y = 0 To (UBound(������_���, 1))
         If ������_���(Y, 0) = �����_��������_���_��_������ Then
            �����������_����_������ = True
            ����(a1, 1) = ������_���(Y, 1)
            ���� = ������_���(Y, 2)
            ����(a1, 4) = ������_���(Y, 3)
            ����(a1, 5) = ������_���(Y, 4)
            ����(a1, 6) = ������_���(Y, 5)
            �����_��� = True
               'If ���_��������_���_��_������ <> ����(a1, 1) Or _
               '   ����_��������_���_��_������ <> ���� Or _
               '   �����_�����_��������_���_��_������ <> ����(a1, 4) Then
                  'UserForm1.Label1.Width = CInt((�����_�����_��������_���_��_������ * 100) / 300)
                  'UserForm1.Label_���� = ����_��������_���_��_������ & " �� " & �����_������_�_�����
                  'UserForm1.Repaint

                  �������_�_����������
               ' End If
                
            Exit For
         End If
       Next Y
              If (���_��������_���_��_������ <> "" Or _
              ����_��������_���_��_������ <> "" Or _
              �����_�����_��������_���_��_������ <> "") And �����_��� = False Then
                 �����������_����_������ = True
                 ����(a1, 1) = ""
                 ���� = ""
                 ����(a1, 4) = ""
                 ����(a1, 5) = ""
                 ����(a1, 6) = ""
                 ����(a1, 7) = ""
                 ����(a1, 8) = ""
                 'UserForm1.Label1.Width = CInt((�����_�����_��������_���_��_������ * 100) / 300)
                 'UserForm1.Label_���� = ����_��������_���_��_������ & " �� " & �����_������_�_�����
                 'UserForm1.Repaint

                 �������_�_����������
                 �����������_����_������ = False
              End If
              �����_��� = False
  Next i

  Application.ScreenUpdating = True '�������� ���������� ���������
  If (Timer - �����) > 60 Then
      Debug.Print (Timer - �����) / 60 & " ���" ' �����  � ���
      Else
      Debug.Print Timer - ����� & " ���" ' �����  � ���
  End If
End Sub



Sub ���������_������������_�_������_�_���������_�_�������()
����� = Timer
Application.ScreenUpdating = False '��������� ���������� ���������
UserForm1.Repaint
DoEvents

'����� ������� � ����������
For i = 1 To ActiveDocument.Tables.Count Step 1
   If ActiveDocument.Tables(i).Range.Columns.Count = 18 Then
      �����_������� = i
      Exit For
   End If
Next i

ActiveDocument.Tables(�����_�������).Range.Cells(1).Select
�����_���_�_������_���� = Selection.Information(wdActiveEndPageNumber) '����� �������� ���
 
UserForm2.TextBox1.Text = ""
UserForm2.TextBox2.Text = ""
UserForm2.TextBox3.Text = ""
UserForm2.TextBox_�1_���.Text = ""
UserForm2.TextBox_�1_���.Text = ""
UserForm2.TextBox_�1_��.Text = ""
UserForm2.TextBox_�2_���.Text = ""
UserForm2.TextBox_�2_���.Text = ""
UserForm2.TextBox_�2_��.Text = ""
UserForm2.TextBox_�3_���.Text = ""
UserForm2.TextBox_�3_���.Text = ""
UserForm2.TextBox_�3_��.Text = ""
UserForm2.TextBox_�4_���.Text = ""
UserForm2.TextBox_�4_���.Text = ""
UserForm2.TextBox_�4_��.Text = ""
UserForm2.TextBox_�5_���.Text = ""
UserForm2.TextBox_�5_���.Text = ""
UserForm2.TextBox_�5_��.Text = ""
UserForm2.TextBox_�6_���.Text = ""
UserForm2.TextBox_�6_���.Text = ""
UserForm2.TextBox_�6_��.Text = ""
UserForm2.TextBox_�7_���.Text = ""
UserForm2.TextBox_�7_���.Text = ""
UserForm2.TextBox_�7_��.Text = ""
UserForm2.TextBox_�8_���.Text = ""
UserForm2.TextBox_�8_���.Text = ""
UserForm2.TextBox_�8_��.Text = ""
UserForm2.TextBox_�9_���.Text = ""
UserForm2.TextBox_�9_���.Text = ""
UserForm2.TextBox_�9_��.Text = ""
UserForm2.TextBox_�10_���.Text = ""
UserForm2.TextBox_�10_���.Text = ""
UserForm2.TextBox_�10_��.Text = ""
UserForm2.TextBox_�11_���.Text = ""
UserForm2.TextBox_�11_���.Text = ""
UserForm2.TextBox_�11_��.Text = ""
UserForm2.TextBox_�12_���.Text = ""
UserForm2.TextBox_�12_���.Text = ""
UserForm2.TextBox_�12_��.Text = ""
UserForm2.TextBox_���_���.Text = ""
UserForm2.TextBox_��_���.Text = ""
UserForm2.TextBox4_��_���.Text = ""
UserForm2.TextBox_��.Text = ""
UserForm2.TextBox5.Text = ""

      If ���_����� = "�10" And ����_����� = "���" Then UserForm2.TextBox_�10_���.Text = UserForm2.TextBox_�10_���.Text + �����_���_����� & ", "
      If ���_����� = "�11" And ����_����� = "���" Then UserForm2.TextBox_�11_���.Text = UserForm2.TextBox_�11_���.Text + �����_���_����� & ", "
      If ���_����� = "�11" And ����_����� = "���" Then UserForm2.TextBox_�11_���.Text = UserForm2.TextBox_�11_���.Text + �����_���_����� & ", "
      If ���_����� = "�12" And ����_����� = "���" Then UserForm2.TextBox_�12_���.Text = UserForm2.TextBox_�12_���.Text + �����_���_����� & ", "
      If ���_����� = "�12" And ����_����� = "���" Then UserForm2.TextBox_�12_���.Text = UserForm2.TextBox_�12_���.Text + �����_���_����� & ", "


If ���������_������_�����_��� = True Then
Else

Call ��������_������_���_����������� '������_���(1,1)

End If '���������_������_�����_��� = True

�����_������_�_����� = ActiveDocument.ComputeStatistics(wdStatisticPages)






'������ �� ��������� � �������� ���������
Erase ������_�����_��� '��������� ������
ReDim ������_�����_���(�����_������_�_����� * 3, 5) '������ ������ ������� ������ � 3 ���� ������ ����� ����� ��������� ���� ��� ��������� ������ �����

�����_��������_��� = �����_���_�_������_���� - 1
  ����������_��� = ""
  For i = �����_���_�_������_���� To �����_������_�_�����
     DoEvents
     
     �����_��������_��� = �����_��������_��� + 1 '����� �������� ���
     Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=i, Name:="" '������� �� ��� �� ������
     'Selection.Range.Select
     'Selection.HomeKey Unit:=wdLine '������� � ������ ������

              
      'Selection.GoTo What:=wdGoToPage, Which:=wdGoToNext, Count:=1, Name:="" '������� �� ��������� ���
      WordBasic.ViewFooterOnly ' ������� ������ ����������
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
      WordBasic.ViewFooterOnly ' ������� ������ ����������
      
       'Application.ScreenUpdating = True '�������� ���������� ���������
       
       'Application.ScreenUpdating = False '�������� ���������� ���������
       
       DoEvents
       UserForm1.Label1.Width = CInt((300 * �����_��������_���) / �����_������_�_�����)
       UserForm1.Label_���� = "���� " & �����_��������_��� & " �� " & �����_������_�_�����
       'If (Timer - �����) > 60 Then
          'UserForm1.Label2 = "����� " & Round((Timer - �����) / 60, 1) & " ���" ' �����  � ���
       '   UserForm1.Label2 = TimeSerial(0, 0, Timer - �����)
       'Else
          'UserForm1.Label2 = "����� " & Round(Timer - �����, 1) & " ���" ' �����  � ���
           UserForm1.Label2 = "�����: " & TimeSerial(0, 0, Timer - �����)
       'End If
       UserForm1.Repaint
      
      �����_���_����� = Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text, Chr(13) & "", "")
      ���_����� = Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 1).Range.Text, Chr(13) & "", "")
      ����_����� = Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 2).Range.Text, Chr(13) & "", "")
      �����_����� = Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 3).Range.Text, Chr(13) & "", "")
      
      �����_���_����� = Replace(�����_���_�����, " ", "")
      ���_����� = Replace(���_�����, " ", "")
      ���_����� = Replace(���_�����, "a", "�")
      ����_����� = Replace(����_�����, " ", "")
      ����_����� = Replace(����_�����, ".", "")
      ����_����� = Replace(����_�����, "���", "���")
      ����_����� = Replace(����_�����, "���", "���")
      �����_����� = Replace(�����_�����, " ", "")
      
      
������_�����_���(i, 0) = �����_���_�����
������_�����_���(i, 1) = ���_�����
������_�����_���(i, 2) = ����_�����
������_�����_���(i, 3) = �����_�����
������_�����_���(i, 4) = CStr(�����_��������_���)
      
If ���������_������_�����_��� = True Then
Else
      If ���_����� = "" Then
            UserForm2.TextBox_���_���.Text = UserForm2.TextBox_���_���.Text + �����_���_����� & ", "
            GoTo �����_���
      End If
      
      If ���_����� = "�1" And ����_����� = "���" Then
            UserForm2.TextBox_�1_���.Text = UserForm2.TextBox_�1_���.Text + �����_���_����� & ", "
            GoTo �����_���
      End If
      If ���_����� = "�1" And ����_����� = "���" Then
            UserForm2.TextBox_�1_���.Text = UserForm2.TextBox_�1_���.Text + �����_���_����� & ", "
            GoTo �����_���
      End If
      If ���_����� = "�1" And ����_����� <> "���" And ����_����� <> "���" Then
            UserForm2.TextBox_�1_��.Text = UserForm2.TextBox_�1_��.Text + �����_���_����� & ", "
            GoTo �����_���
      End If
      If ���_����� = "�2" And ����_����� = "���" Then
            UserForm2.TextBox_�2_���.Text = UserForm2.TextBox_�2_���.Text + �����_���_����� & ", "
            GoTo �����_���
      End If
      If ���_����� = "�2" And ����_����� = "���" Then
            UserForm2.TextBox_�2_���.Text = UserForm2.TextBox_�2_���.Text + �����_���_����� & ", "
            GoTo �����_���
      End If
            If ���_����� = "�2" And ����_����� <> "���" And ����_����� <> "���" Then
            UserForm2.TextBox_�2_��.Text = UserForm2.TextBox_�2_��.Text + �����_���_����� & ", "
            GoTo �����_���
      End If
      If ���_����� = "�3" And ����_����� = "���" Then
            UserForm2.TextBox_�3_���.Text = UserForm2.TextBox_�3_���.Text + �����_���_����� & ", "
            GoTo �����_���
      End If
      If ���_����� = "�3" And ����_����� = "���" Then
            UserForm2.TextBox_�3_���.Text = UserForm2.TextBox_�3_���.Text + �����_���_����� & ", "
            GoTo �����_���
      End If
            If ���_����� = "�3" And ����_����� <> "���" And ����_����� <> "���" Then
            UserForm2.TextBox_�3_��.Text = UserForm2.TextBox_�3_��.Text + �����_���_����� & ", "
            GoTo �����_���
      End If
      If ���_����� = "�4" And ����_����� = "���" Then
            UserForm2.TextBox_�4_���.Text = UserForm2.TextBox_�4_���.Text + �����_���_����� & ", "
            GoTo �����_���
      End If
      If ���_����� = "�4" And ����_����� = "���" Then
            UserForm2.TextBox_�4_���.Text = UserForm2.TextBox_�4_���.Text + �����_���_����� & ", "
            GoTo �����_���
      End If
            If ���_����� = "�4" And ����_����� <> "���" And ����_����� <> "���" Then
            UserForm2.TextBox_�4_��.Text = UserForm2.TextBox_�4_��.Text + �����_���_����� & ", "
            GoTo �����_���
      End If
      If ���_����� = "�5" And ����_����� = "���" Then
            UserForm2.TextBox_�5_���.Text = UserForm2.TextBox_�5_���.Text + �����_���_����� & ", "
            GoTo �����_���
      End If
      If ���_����� = "�5" And ����_����� = "���" Then
            UserForm2.TextBox_�5_���.Text = UserForm2.TextBox_�5_���.Text + �����_���_����� & ", "
            GoTo �����_���
      End If
            If ���_����� = "�5" And ����_����� <> "���" And ����_����� <> "���" Then
            UserForm2.TextBox_�5_��.Text = UserForm2.TextBox_�5_��.Text + �����_���_����� & ", "
            GoTo �����_���
      End If
      If ���_����� = "�6" And ����_����� = "���" Then
            UserForm2.TextBox_�6_���.Text = UserForm2.TextBox_�6_���.Text + �����_���_����� & ", "
            GoTo �����_���
      End If
      If ���_����� = "�6" And ����_����� = "���" Then
            UserForm2.TextBox_�6_���.Text = UserForm2.TextBox_�6_���.Text + �����_���_����� & ", "
            GoTo �����_���
      End If
            If ���_����� = "�6" And ����_����� <> "���" And ����_����� <> "���" Then
            UserForm2.TextBox_�6_��.Text = UserForm2.TextBox_�6_��.Text + �����_���_����� & ", "
            GoTo �����_���
      End If
      If ���_����� = "�7" And ����_����� = "���" Then
            UserForm2.TextBox_�7_���.Text = UserForm2.TextBox_�7_���.Text + �����_���_����� & ", "
            GoTo �����_���
      End If
      If ���_����� = "�7" And ����_����� = "���" Then
            UserForm2.TextBox_�7_���.Text = UserForm2.TextBox_�7_���.Text + �����_���_����� & ", "
            GoTo �����_���
      End If
            If ���_����� = "�7" And ����_����� <> "���" And ����_����� <> "���" Then
            UserForm2.TextBox_�7_��.Text = UserForm2.TextBox_�7_��.Text + �����_���_����� & ", "
            GoTo �����_���
      End If
      If ���_����� = "�8" And ����_����� = "���" Then
            UserForm2.TextBox_�8_���.Text = UserForm2.TextBox_�8_���.Text + �����_���_����� & ", "
            GoTo �����_���
      End If
      If ���_����� = "�8" And ����_����� = "���" Then
            UserForm2.TextBox_�8_���.Text = UserForm2.TextBox_�8_���.Text + �����_���_����� & ", "
            GoTo �����_���
      End If
            If ���_����� = "�8" And ����_����� <> "���" And ����_����� <> "���" Then
            UserForm2.TextBox_�8_��.Text = UserForm2.TextBox_�8_��.Text + �����_���_����� & ", "
            GoTo �����_���
      End If
      If ���_����� = "�9" And ����_����� = "���" Then
            UserForm2.TextBox_�9_���.Text = UserForm2.TextBox_�9_���.Text + �����_���_����� & ", "
            GoTo �����_���
      End If
      If ���_����� = "�9" And ����_����� = "���" Then
            UserForm2.TextBox_�9_���.Text = UserForm2.TextBox_�9_���.Text + �����_���_����� & ", "
            GoTo �����_���
      End If
            If ���_����� = "�9" And ����_����� <> "���" And ����_����� <> "���" Then
            UserForm2.TextBox_�9_��.Text = UserForm2.TextBox_�9_��.Text + �����_���_����� & ", "
            GoTo �����_���
      End If
      If ���_����� = "�10" And ����_����� = "���" Then
            UserForm2.TextBox_�10_���.Text = UserForm2.TextBox_�10_���.Text + �����_���_����� & ", "
            GoTo �����_���
      End If
      If ���_����� = "�10" And ����_����� = "���" Then
            UserForm2.TextBox_�10_���.Text = UserForm2.TextBox_�10_���.Text + �����_���_����� & ", "
            GoTo �����_���
      End If
            If ���_����� = "�10" And ����_����� <> "���" And ����_����� <> "���" Then
            UserForm2.TextBox_�10_��.Text = UserForm2.TextBox_�10_��.Text + �����_���_����� & ", "
            GoTo �����_���
      End If
      If ���_����� = "�11" And ����_����� = "���" Then
            UserForm2.TextBox_�11_���.Text = UserForm2.TextBox_�11_���.Text + �����_���_����� & ", "
            GoTo �����_���
      End If
      If ���_����� = "�11" And ����_����� = "���" Then
            UserForm2.TextBox_�11_���.Text = UserForm2.TextBox_�11_���.Text + �����_���_����� & ", "
            GoTo �����_���
      End If
            If ���_����� = "�11" And ����_����� <> "���" And ����_����� <> "���" Then
            UserForm2.TextBox_�11_��.Text = UserForm2.TextBox_�11_��.Text + �����_���_����� & ", "
            GoTo �����_���
      End If
      If ���_����� = "�12" And ����_����� = "���" Then
            UserForm2.TextBox_�12_���.Text = UserForm2.TextBox_�12_���.Text + �����_���_����� & ", "
            GoTo �����_���
      End If
      If ���_����� = "�12" And ����_����� = "���" Then
            UserForm2.TextBox_�12_���.Text = UserForm2.TextBox_�12_���.Text + �����_���_����� & ", "
            GoTo �����_���
      End If
            If ���_����� = "�12" And ����_����� <> "���" And ����_����� <> "���" Then
            UserForm2.TextBox_�12_��.Text = UserForm2.TextBox_�12_��.Text + �����_���_����� & ", "
            GoTo �����_���
      End If
      

      If ����_����� = "���" Then
            UserForm2.TextBox_��_���.Text = UserForm2.TextBox_��_���.Text + �����_���_����� & ", "
            GoTo �����_���
      End If
      If ����_����� = "���" Then
            UserForm2.TextBox4_��_���.Text = UserForm2.TextBox4_��_���.Text + �����_���_����� & ", "
            GoTo �����_���
      End If
      If ����_����� <> "���" And ����_����� <> "���" Then
            UserForm2.TextBox_��.Text = UserForm2.TextBox_��.Text + �����_���_����� & ", "
            GoTo �����_���
      End If
      
      UserForm2.TextBox5.Text = UserForm2.TextBox5.Text + �����_���_����� & ", "
      
�����_���:

End If '���������_������_�����_��� = True
      
     ' '������� ����
     ' For u = 0 To UBound(������_���, 1) Step 1
     '   ������_���(u, 1) = Replace(������_���(u, 1), "a", "�")
     '   For Y = 0 To UBound(������_���, 2) Step 1
     '      ������_���(u, Y) = Replace(������_���(u, Y), " ", "")
     '   Next Y
     ' Next u
      
      '������� ������ ���������� ���
     ' For u = 0 To UBound(������_���, 1) Step 1
     '    If ������_���(u, 0) = �����_���_����� Then
     '      If ������_���(u, 1) <> ���_����� Or _
    '          Replace(������_���(u, 2), ".", "") <> ����_����� Or _
    '          ������_���(u, 3) <> �����_����� Then
    '          ����������_��� = ����������_��� & �����_���_����� & ", "
    '          ��������_����������_��� = ��������_����������_��� & �����_��������_��� & ", "
    '          Exit For
    '       End If
    '     End If
    '  Next u
      '������_�������(j) = Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text, '��������� � ������ � ���� � ������ �����������
      'Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text, Chr(13) & "", "")
      '������_�������(j) = Replace(������_�������(j), "", "")
      '������_�������(j) = Replace(������_�������(j), Chr(13), "")
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������

  Next i
  
If ���������_������_�����_��� = True Then

Else
  
'��������� ����������� ��� � ������_��� (������ ��� ���)
For i = 0 To UBound(������_�����_���, 1)
   If CStr(������_�����_���(i, 0)) <> "" Then
     For r = 0 To UBound(������_���, 1)
        If CStr(������_�����_���(i, 0)) = CStr(������_���(r, 0)) Then
           GoTo �����_�����  '���� �����
        End If
        If r = UBound(������_���, 1) Then '���� ��� ����������� � �� �����
           For j = 0 To UBound(������_���, 1)
              If CStr(������_���(j, 0)) = "" Then
                 ������_���(j, 0) = CStr(������_�����_���(i, 0))
                 GoTo �����_�����  '���� �����
              End If
           Next j
        End If
     Next r
   
�����_�����:
    End If
Next i
'

      '������� ����
      For u = 0 To UBound(������_���, 1) Step 1
        ������_���(u, 1) = Replace(������_���(u, 1), "a", "�")
        For Y = 0 To UBound(������_���, 2) Step 1
           ������_���(u, Y) = Replace(������_���(u, Y), " ", "")
        Next Y
      Next u
      
'���������� �� ������������ ��� �� ������ � ��  ����� ������� ���. (����������_���)
For i = 0 To UBound(������_�����_���, 1)
      '������� ������ ���������� ���
   If CStr(������_�����_���(i, 0)) <> "" Then
      For u = 0 To UBound(������_���, 1) Step 1
         If CStr(������_���(u, 0)) = CStr(������_�����_���(i, 0)) Then
           If CStr(������_���(u, 1)) <> CStr(������_�����_���(i, 1)) Or _
              Replace(������_���(u, 2), ".", "") <> ������_�����_���(i, 2) Or _
              CStr(������_���(u, 3)) <> CStr(������_�����_���(i, 3)) Then
              If CStr(������_���(u, 1)) = "" Then
                ��� = "���"
              Else
                ��� = CStr(������_���(u, 1))
              End If
              If CStr(������_�����_���(i, 1)) = "" Then
                ���_����� = "���"
              Else
                ���_����� = CStr(������_�����_���(i, 1))
              End If
              ����������_��� = ����������_��� & ������_�����_���(i, 0) & "-" & ��� & "/" & ���_����� & ", "
              ��������_����������_��� = ��������_����������_��� & ������_�����_���(i, 4) & ", "
              Exit For
           End If
         End If
      Next u
  End If
Next i
  
If ����������_��� <> "" Then
  UserForm2.TextBox1.Text = Mid(����������_���, 1, Len(����������_���) - 2)
  UserForm2.TextBox2.Text = Mid(��������_����������_���, 1, Len(��������_����������_���) - 2)
End If


  Application.ScreenUpdating = True '�������� ���������� ���������
  UserForm1.Hide

  If ����������_��� = "" Then
    ��������� = MsgBox("���������� ������������ �� �������", vbOKOnly + vbInformation, "�������� ������������")
  Else
    UserForm2.Show
  End If
End If '���������_������_�����_��� = True

  If (Timer - �����) > 60 Then
      Debug.Print (Timer - �����) / 60 & " ���" ' �����  � ���
      Else
      Debug.Print Timer - ����� & " ���" ' �����  � ���
  End If


End Sub



Sub �������_�������()

Application.ScreenUpdating = False '��������� ���������� ���������
Dim ������_�����_������_�_���������(), ������_����������_���()

�����_������_�_����� = ActiveDocument.ComputeStatistics(wdStatisticPages)
�����_��������_�_����� = ActiveDocument.Sections.Count '���������� ��������
'�����_��������_��� = Selection.Information(wdActiveEndPageNumber) '����� �������� ���
'Selection.Information (wdActiveEndSectionNumber) '����� ��������� �������
'ActiveDocument.Sections.Count'���������� ��������
'Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=�����_������_�_����� - 2, Name:="" '������� �� ��� �� ������
'Selection.GoTo What:=wdGoToSection, Which:=wdGoToFirst, Count:=5, Name:="" '������� � ������

'������� ������ �������
    ��������� = MsgBox("�������� ������ ������� �� ����� TXT: Yes-""��"", No-""������� ����� ������"", Cancel-""�����"" ", vbYesNoCancel + vbQuestion + vbDefaultButton1, "������ �� ������ �������")
    If ��������� = vbCancel Then Exit Sub
    'If ��������� = vbYes Then ���� = "���."
    If ��������� = vbNo Then Call ��������_������_���_�_����


���� = ActiveDocument.Path
��� = ActiveDocument.Name

'�������� ���������� �� ����
If Dir(���� & "\" & ��� & "_��������.txt") = "" Then
  ��������� = MsgBox("���� �� ������. ������� ����� OK-""��"",Cancel-""�����"" ", vbOKCancel + vbQuestion + vbDefaultButton1, "������")
  If ��������� = vbCancel Then Exit Sub
  If ��������� = vbOK Then Call ��������_������_���_�_����
End If

Open ���� & "\" & ��� & "_��������.txt" For Input As #1  '������� ��� ������ "Input"
Line Input #1, �����_��_txt '������ ������
Close #1 '������� ��������

������ = Split(�����_��_txt, ",") '�������� ������
ReDim ������_�������(UBound(������, 1)) '������ ������ �������
ReDim ������_�����_������_�_���������(UBound(������_�������, 1), UBound(������_�������, 1)) '������ ������ �������
ReDim ������_����������_���(�����_������_�_�����, 2)



'����������� �������� ����� ������� �������
For i = 0 To UBound(������, 1)
   ������_�������(i) = ������(i)
Next i

����� = 0

'������ �� ��������� � ���� ���������
  For i = 0 To UBound(������_�������, 1) Step 1
      '�������� ������������� ��� � ������
      If (i = UBound(������_�������, 1) - 1) And CInt(������_�������(UBound(������_�������, 1) - 1 - 1)) = ������_�������(UBound(������_�������, 1) - 1 - 1) Then
        ������_�����_������_�_���������(�����, 0) = ������_�������(i + 1) '����� �� ����� �������
        ������_�����_������_�_���������(�����, 1) = CStr(i + 1 + 1) '����� �������� ��������
        ����� = ����� + 1
      End If
     'Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=i, Name:="" '������� �� ��� �� ������
      '�������� �� ����� ��� � �����
      If CInt(������_�������(i)) <> ������_�������(i) Then
        'ReDim Preserve ������_�����_������_�_���������(UBound(������_�����_������_�_���������, 1) + 1)

        ������_�����_������_�_���������(�����, 0) = ������_�������(i) '����� �� ����� �������
        ������_�����_������_�_���������(�����, 1) = CStr(i + 1) '����� �������� ��������
        ����� = ����� + 1
      End If
  Next i

'����������� ������ �������� �������
'ActiveDocument.Range.Tables.Count


'Set tblSel = Selection.Tables(1)
'ingStart = tblSel.Range.Start
'����� ������� � ����������
For i = 1 To ActiveDocument.Tables.Count Step 1
   If ActiveDocument.Tables(i).Range.Columns.Count = 18 Then
      �����_������� = i
      Exit For
   End If
Next i

ActiveDocument.Tables(�����_�������).Range.Cells(1).Select
�����_���_�_������_���� = Selection.Information(wdActiveEndPageNumber) '����� �������� ���

For i = UBound(������_�����_������_�_���������, 1) To 0 Step -1
   If ������_�����_������_�_���������(i, 1) <> Empty Then
     If i <> 0 Then
       If CInt(������_�����_������_�_���������(i, 0)) <> (������_�����_������_�_���������(i, 0)) And _
          CInt(������_�����_������_�_���������(i - 1, 0)) <> (������_�����_������_�_���������(i - 1, 0)) And _
          (������_�����_������_�_���������(i, 0)) - (������_�����_������_�_���������(i - 1, 0)) <= 0.1 Then
          GoTo ���������_����
       End If
     End If
      ��������_��� = ������_�����_������_�_���������(i, 1) - 1
      If i = 0 Then
              ���������_��� = �����_���_�_������_����
        Else: ���������_��� = ������_�����_������_�_���������(i - 1, 1) + 1
      End If
      'If ��������_��� - ���������_��� = 1 Then GoTo ���������_����
      
     
      
      ������_���_����� = ������_�������(��������_��� - 1)
      ������_���_����� = ������_�������(���������_��� - 1)
            
     ' If (CInt(������_���_�����) = ������_���_�����) And _
     '    (CInt(������_���_�����) <> ������_���_�����) And _
     '    ������_���_����� - ������_���_����� < 1 Then
     ' GoTo ���������_����
     '  End If
          
      'End If
      'If CInt(������_�������(��������_��� - 2)) = (������_�������(��������_��� - 2)) Then
      '    ������_�����_��� = ��������_��� - 1
      'End If
      
      
      ������_�����_��� = ��������_���
      Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=������_�����_���, Name:="" '������� �� ��� �� ������
      
       DoEvents
       UserForm1.Label1.Width = CInt((300 * (�����_������_�_����� - �����_��������_���)) / �����_������_�_�����)
       UserForm1.Label_���� = "���� " & (�����_������_�_����� - �����_��������_���) & " �� " & �����_������_�_�����
       UserForm1.Repaint
      
      
      For Y = ��������_��� To (���������_��� + 1) Step -1
         'If Y = (���������_��� + 1) Then Exit For
    'ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader '������� � ������� ����������
    'Selection.HeaderFooter.LinkToPrevious = Not Selection.HeaderFooter.LinkToPrevious ' ��������� ��� � ���������� �������
    'ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument '������� ����������
         If Selection.Information(wdActiveEndSectionNumber) = 3 Then  '����� ��������� �������
            ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader '������� � ������� ����������
            Selection.HeaderFooter.LinkToPrevious = False ' ��������� ��� � ���������� �������
            ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument '������� ����������
            Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=������_�����_���, Name:="" '������� �� ��� �� ������
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
                 
          WordBasic.ViewFooterOnly ' ������� ������ ����������
          
          On Error GoTo ���������_���� '���������� ������
          �����_�����_��� = Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text, Chr(13) & "", "") '��������� � ������ � ���� � ������ �����������
          On Error GoTo 0
          
            For h = 0 To UBound(������_�������, 1) Step 1
                If �����_�����_��� = ������_�������(h) Then
                   �����_��� = h + 1
                End If
            Next h
          
          ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������

         If �����_��� < ���������_��� Then
            GoTo ���������_����
         End If
         Selection.Delete Unit:=wdCharacter, Count:=1
         
         Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=������_�����_���, Name:="" '������� �� ��� �� ������

      
         '�������� �����������
          WordBasic.ViewFooterOnly ' ������� ������ ����������
          ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
          WordBasic.ViewFooterOnly ' ������� ������ ����������
          �����_�����_��� = Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text, Chr(13) & "", "") '��������� � ������ � ���� � ������ �����������
          
          '������� �����������
          Selection.HeaderFooter.Range.Tables(1).Cell(2, 1).Range.Text = "" '������� � �����������
          Selection.HeaderFooter.Range.Tables(1).Cell(2, 2).Range.Text = "" '������� � �����������
          Selection.HeaderFooter.Range.Tables(1).Cell(2, 3).Range.Text = "" '������� � �����������
          Selection.HeaderFooter.Range.Tables(1).Cell(2, 4).Range.Text = "" '������� � �����������
          Selection.HeaderFooter.Range.Tables(1).Cell(2, 5).Range.Text = "" '������� � �����������
          '������_����������_���()
          
          '�����_��������_��� = Y 'Selection.Information(wdActiveEndPageNumber) '����� �������� ���
        
       �����_��������_��� = Selection.Information(wdActiveEndPageNumber) '����� �������� ���
         
         If �����_��������_��� < �����_���_�_������_���� Then
            GoTo ���������_����
         End If
         
       DoEvents
       UserForm1.Label1.Width = CInt((300 * (�����_������_�_����� - �����_��������_���)) / �����_������_�_�����)
       UserForm1.Label_���� = "���� " & (�����_������_�_����� - �����_��������_���) & " �� " & �����_������_�_�����
       UserForm1.Repaint


 
 
       'If �����_�����_��� <> ������_�������(������_�����_��� - 1) Then '(Selection.Information(wdActiveEndPageNumber)) Then
       If Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text, Chr(13) & "", "") <> ������_�������(Selection.Information(wdActiveEndPageNumber) - 1) And _
        ������_�������(Selection.Information(wdActiveEndPageNumber) - 1) <> 1 Then
            'Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=�����_��������_��� + 1, Name:="" '������� �� ��� �� ������
           'WordBasic.ViewFooterOnly ' ������� ������ ����������
            If Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text, Chr(13) & "", "") <> ������_�������(Selection.Information(wdActiveEndPageNumber) - 1) Then
               With Selection.HeaderFooter.PageNumbers  ' ������ ����� ��� �������
                 .NumberStyle = wdPageNumberStyleArabic
                 .HeadingLevelForChapter = 0
                 .IncludeChapterNumber = False
                 .ChapterPageSeparator = wdSeparatorHyphen
                 .RestartNumberingAtSection = True
                 .StartingNumber = ������_�������(���������_��� - 1)
              End With
            End If
           
           'Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=������_�����_���, Name:="" '������� �� ��� �� ������
           'WordBasic.ViewFooterOnly ' ������� ������ ����������
           
           If Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text, Chr(13) & "", "") <> ������_�������(Selection.Information(wdActiveEndPageNumber) - 1) Then
              With Selection.HeaderFooter.PageNumbers  ' ������ ����� ��� �������
                 .NumberStyle = wdPageNumberStyleArabic
                 .HeadingLevelForChapter = 0
                 .IncludeChapterNumber = False
                 .ChapterPageSeparator = wdSeparatorHyphen
                 .RestartNumberingAtSection = True
                 .StartingNumber = 1
              End With
           End If

            If Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text, Chr(13) & "", "") <> ������_�������(Selection.Information(wdActiveEndPageNumber) - 1) Then
             With Selection.HeaderFooter.PageNumbers  ' ���������� �������
               .NumberStyle = wdPageNumberStyleArabic
               .HeadingLevelForChapter = 0
               .IncludeChapterNumber = False
               .ChapterPageSeparator = wdSeparatorHyphen
               .RestartNumberingAtSection = False
               .StartingNumber = 0
             End With
            End If
      End If
           ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
           
           ������_�����_��� = ������_�����_��� - 1
         
      Next Y
   End If
���������_����:
ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������

Next i

'�������� ������ �� "�� ��������� ��������"
Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=�����_���_�_������_����, Name:="" '������� �� ��� �� ������

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

    'ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader '������� � ������� ����������
    'Selection.HeaderFooter.LinkToPrevious = Not Selection.HeaderFooter.LinkToPrevious ' ��������� ��� � ���������� �������
    'ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument '������� ����������
Application.ScreenUpdating = True '��������� ���������� ���������

End Sub






Sub ��������_��������_�_������_����_������()
  'UserForm1.Repaint
  UserForm1.Label1.Width = 0
  'DoEvents
  �����_��������_�_������_����_������ = True
  UserForm1.Show
  �����_��������_�_������_����_������ = False
  �����������_����_������ = False
  'UserForm1.Hide
End Sub


Sub ��������_������_���_�_����()
  'UserForm1.Repaint
   UserForm1.Label1.Width = 0
   UserForm1.Label_���� = "����������"
  'DoEvents
  �����_������_���_�_���� = True
  On Error Resume Next
  UserForm1.Show
  On Error GoTo 0  '����� �� ���������� ������
  �����_������_���_�_���� = False
  'UserForm1.Hide
End Sub

Sub ��������_���������_������������_�_������_�_���������_�_�������()
    '��������� = MsgBox("��������� ����� ��� ������ ������� ���� � ������������ �������� ������: Yes-""�����"", No-""������� �����"", Cancel-""�����"" ", vbYesNoCancel + vbQuestion + vbDefaultButton1, "������ �� ������ �������")
    'If ��������� = vbCancel Then Exit Sub
    'If ��������� = vbYes Then ���� = "���."
    'If ��������� = vbNo Then
    '  UserForm2.Show
    '  Exit Sub
    'End If
  'UserForm1.Repaint
   UserForm1.Label1.Width = 0
   UserForm1.Label_���� = "����������"
  'DoEvents
  �����_���������_������������_�_������_�_���������_�_������� = True
  '��_������_������_����� = True
  UserForm1.Show
  �����_���������_������������_�_������_�_���������_�_������� = False
  'UserForm1.Hide
End Sub

Sub ��������_�������_�������()
  'UserForm1.Repaint
   UserForm1.Label1.Width = 0
   UserForm1.Label_���� = "����������"
  'DoEvents
  �����_�������_������� = True
  UserForm1.Show
  �����_�������_������� = False
  'UserForm1.Hide
End Sub
'�������_�������

Sub ����������()
  UserForm3.Show
End Sub



Sub ������_�_�������()
Attribute ������_�_�������.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.������_�_�������"
'
' ������_�_������� ������
'
'
    If Selection.Range.HighlightColorIndex = wdBrightGreen Then
       Selection.Range.HighlightColorIndex = wdNoHighlight
    Else
       Selection.Range.HighlightColorIndex = wdBrightGreen
    End If
End Sub
Sub ������_�_�������()
Attribute ������_�_�������.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.������_�_�������"
'
' ������_�_������� ������
'
'
    If Selection.Range.HighlightColorIndex = wdRed Then
       Selection.Range.HighlightColorIndex = wdNoHighlight
    Else
       Selection.Range.HighlightColorIndex = wdRed
    End If
    'Options.DefaultHighlightColorIndex = wdRed
End Sub
Sub ������_�_������_1()
Attribute ������_�_������_1.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.������_�_������"
'
' ������_�_������_1 ������
'
'
    If Selection.Range.HighlightColorIndex = wdYellow Then
       Selection.Range.HighlightColorIndex = wdNoHighlight
    Else
       Selection.Range.HighlightColorIndex = wdYellow
    End If
    'Selection.Range.HighlightColorIndex = wdYellow
End Sub

Sub ������_�_�������()
'
' ������_�_������� ������
'
'
    If Selection.Range.HighlightColorIndex = wdTurquoise Then
       Selection.Range.HighlightColorIndex = wdNoHighlight
    Else
       Selection.Range.HighlightColorIndex = wdTurquoise
    End If
    'Selection.Range.HighlightColorIndex = wdTurquoise
End Sub

Sub ������_�_�����()
'
' ������_�_����� ������
'
'
    If Selection.Range.HighlightColorIndex = wdGray25 Then
       Selection.Range.HighlightColorIndex = wdNoHighlight
    Else
       Selection.Range.HighlightColorIndex = wdGray25
    End If
End Sub

Sub ��������_������_�����_��_�������()
  'UserForm1.Repaint
   UserForm1.Label1.Width = 0
   UserForm1.Label_���� = "����������"
  'DoEvents
  �����_������_�����_��_������� = True
  UserForm1.Show
  �����_������_�����_��_������� = False
  'UserForm1.Hide
End Sub





Sub �������_������_������()

'����������� ������ �������� �������
Set tblSel = Selection.Tables(1)
ingStart = tblSel.Range.Start
For i = 1 To ActiveDocument.Tables.Count Step 1
 If ActiveDocument.Tables(i).Range.Start = ingStart Then
    ingTbIndex = i
    Exit For
 End If
 Next i
 
'������� ������ ���� 2 � 4 ������ �����
For i = 1 To ActiveDocument.Tables(ingTbIndex).Rows.Count Step 1 '������ �� �������
   'ActiveDocument.Tables(1).Rows.Select
   'For r = 1 To ActiveDocument.Tables(�1).Columns.Count Step 1
   'If Err.Number = 5941 Then GoTo �������
   'On Error GoTo �������  '���������� ������
    On Error Resume Next
    f = ActiveDocument.Tables(ingTbIndex).Cell(i, 2).Range.Text '���� �������� Err.Number ��������� ������
    If Err.Number = 5941 Then
      GoTo �������
    End If
    f = ActiveDocument.Tables(ingTbIndex).Cell(i, 4).Range.Text '���� �������� Err.Number ��������� ������
    If Err.Number = 5941 Then
      GoTo �������
    End If
        If Replace(ActiveDocument.Tables(ingTbIndex).Cell(i, 2).Range.Text, Chr(13) & "", "") = "" And Replace(ActiveDocument.Tables(ingTbIndex).Cell(i, 4).Range.Text, Chr(13) & "", "") = "" And Err.Number <> 5941 Then
          'On Error GoTo �������  '���������� ������
          ActiveDocument.Tables(ingTbIndex).Cell(i, 1).Select
          'ActiveDocument.Tables(ingTbIndex).Rows(i).Select
          Selection.Rows.Delete
        End If
�������:
    Err.Clear

    DoEvents
Next i
On Error GoTo 0
End Sub

Sub ������_�����_��_�������()
  
t = Timer

Application.ScreenUpdating = False '��������� ���������� ���������
����������_������ = ActiveDocument.Tables.Count

For i = 1 To ActiveDocument.Tables.Count Step 1
       '�����_��������_��� = Selection.Information(wdActiveEndPageNumber) '����� �������� ���
       ����������_����� = ActiveDocument.Tables(i).Range.Rows.Count
       DoEvents
       UserForm1.Label1.Width = CInt((300 * i) / ����������_������)
       UserForm1.Label_���� = "������� " & i & " �� " & ����������_������ & " (������ " & g & " �� " & ����������_����� & " )"
       'If (Timer - t) > 60 Then
       '   UserForm1.Label2 = "����� " & Round((Timer - t) / 60, 1) & " ���" ' �����  � ���
       'Else
       '   UserForm1.Label2 = "����� " & Round(Timer - t, 1) & " ���" ' �����  � ���
       'End If
       UserForm1.Label2 = "�����: " & TimeSerial(0, 0, Timer - t)
       
       UserForm1.Repaint
 
 If ActiveDocument.Tables(i).Range.Columns.Count = 18 Then
    For g = 1 To ActiveDocument.Tables(i).Range.Rows.Count Step 1
        
        DoEvents
        UserForm1.Label2 = "�����: " & TimeSerial(0, 0, Timer - t)
        UserForm1.Label_���� = "������� " & i & " �� " & ����������_������ & " (������ " & g & " �� " & ����������_����� & " )"
        UserForm1.Repaint
        
        If ActiveDocument.Tables(i).Cell(g, 5).Range <> Chr(13) & "" Or _
           ActiveDocument.Tables(i).Cell(g, 6).Range <> Chr(13) & "" Or _
           ActiveDocument.Tables(i).Cell(g, 7).Range <> Chr(13) & "" Then
           'Replace(�������, Chr(13) & "", "")
           ��� = ActiveDocument.Tables(i).Cell(g, 5).Range
           ����� = ActiveDocument.Tables(i).Cell(g, 6).Range
           ��� = ActiveDocument.Tables(i).Cell(g, 7).Range
           ��� = Replace(���, ".", ",")
           ����� = Replace(�����, ".", ",")
           ��� = Replace(���, ".", ",")
           ��� = Replace(���, Chr(13) & "", "")
           ����� = Replace(�����, Chr(13) & "", "")
           ��� = Replace(���, Chr(13) & "", "")
           ��� = Replace(���, " ", "")
           ����� = Replace(�����, " ", "")
           ��� = Replace(���, " ", "")
        
           ActiveDocument.Tables(i).Cell(g, 5).Range = ���
           ActiveDocument.Tables(i).Cell(g, 6).Range = �����
           ActiveDocument.Tables(i).Cell(g, 7).Range = ���
        End If
    Next g
 End If
 Next i
 
     If (Timer - t) > 60 Then
      Debug.Print (Timer - t) / 60 & " ���" ' �����  � ���
      Else
      Debug.Print Timer - t & " ���" ' �����  � ���
    End If
    
 
Application.ScreenUpdating = True '�������� ���������� ���������

End Sub


Sub ������_�_�����������_������_���������_���������_����()


Sub ������_�_�����������_������_���������_���������()

 Dim ������_���_���()
 
  ����� = Timer
  
  DoEvents
  
'����� ������� � ����������
For i = 1 To ActiveDocument.Tables.Count Step 1
   If ActiveDocument.Tables(i).Range.Columns.Count = 18 Then
      �����_������� = i
      Exit For
   End If
Next i

ActiveDocument.Tables(�����_�������).Range.Cells(1).Select
�����_���_�_������_���� = Selection.Information(wdActiveEndPageNumber) '����� �������� ���

Application.ScreenUpdating = False '��������� ���������� ���������
  'UserForm1.Label1.Width = 0
  'UserForm1.Show
�����_������_�_����� = ActiveDocument.ComputeStatistics(wdStatisticPages)

   'Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=�����_��������_���, Name:="" '������� �� ��� �� ������


'�������� ������ � ��������� ������� �� ��������� ����
Call ��������_������_���_�����������  '������_���()



'Erase ���� '��������� ������
'ReDim ����(.Rows.Count * 2, .Rows.Count * 2) '������ ������ �������
'For i = 0 To UBound(������_���, 1)
'     ����(i, 1) = ������_���(i, 1)
'     ����(i, 2) = ������_���(i, 2)
'     ����(i, 4) = ������_���(i, 3)
'     ����(i, 5) = ������_���(i, 4)
'     ����(i, 6) = ������_���(i, 5)
'Next i

'���������� ����� ���������� �������
   With ActiveDocument.Range.Tables(ActiveDocument.Range.Tables.Count)
      For i = 4 To .Rows.Count  ' ��������� ��� �� ��������� ���� ��� ������� ��� (�1, �2 ...)
         If .Cell(i, 1).Range.Text <> Chr(13) & "" Then
            ���_������� = Replace(.Cell(i, 1).Range.Text, Chr(13) & "", "")
            '���� = "���."

         End If
      Next i
   End With
   
   
   
'������� ��������� ���. � ������ �����������
For i = �����_���_�_������_���� To �����_������_�_����� - 1



       Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=i, Name:="" '������� �� ��� �� ������
       

       �����_��������_��� = Selection.Information(wdActiveEndPageNumber) '����� �������� ���
       DoEvents
       UserForm1.Label1.Width = CInt((300 * �����_��������_���) / �����_������_�_�����)
       UserForm1.Label_���� = "���� " & �����_��������_��� & " �� " & �����_������_�_�����
       'If (Timer - t) > 60 Then
       '   UserForm1.Label2 = "����� " & Round((Timer - t) / 60, 1) & " ���" ' �����  � ���
       'Else
       '   UserForm1.Label2 = "����� " & Round(Timer - t, 1) & " ���" ' �����  � ���
       'End If
       UserForm1.Label2 = "�����: " & TimeSerial(0, 0, Timer - �����)
       UserForm1.Repaint
       
       'For g = 0 To UBound(������_���, 1)

       
       WordBasic.ViewFooterOnly ' ������� ������ ����������
       ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
       WordBasic.ViewFooterOnly ' ������� ������ ����������
       
       '����� ���� �� �� ������� ��� ��������� ���.
       For g = 0 To UBound(������_���, 1)
           If CStr(������_���(g, 1)) = CStr(���_�������) And CStr(������_���(g, 0)) = Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text, Chr(13) & "", "") Then
              
              ���� = ������_���(g, 2)
              �����������_������_��_�������_��� = True
              Call �������_�_����������
              �����������_������_��_�������_��� = False
              
       '       Selection.HeaderFooter.Range.Tables(1).Cell(2, 1).Range.Text = ������_���(g, 1)
       '       Selection.HeaderFooter.Range.Tables(1).Cell(2, 2).Range.Text = ������_���(g, 2)
       '       Selection.HeaderFooter.Range.Tables(1).Cell(2, 3).Range.Text = ������_���(g, 3)
       '       Selection.HeaderFooter.Range.Tables(1).Cell(2, 4).Range.Text = ������_���(g, 4)
       '       Selection.HeaderFooter.Range.Tables(1).Cell(2, 5).Range.Text = ������_���(g, 5)
       '       ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
              Exit For
             '�����_��������_���_��_������ = Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text, Chr(13) & "", "")
             '���_��������_���_��_������ = Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 1).Range.Text, Chr(13) & "", "")
              '����_��������_���_��_������ = Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 2).Range.Text, Chr(13) & "", "")
             '�����_�����_��������_���_��_������ = Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 3).Range.Text, Chr(13) & "", "")
           End If
       Next g

    

Next i

  ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
  
  Application.ScreenUpdating = True '�������� ���������� ���������
  If (Timer - �����) > 60 Then
      Debug.Print (Timer - �����) / 60 & " ���" ' �����  � ���
      Else
      Debug.Print Timer - ����� & " ���" ' �����  � ���
  End If
  
End Sub


Sub ��������_������_�_�����������_������_���������_���������()
  'UserForm1.Repaint
   UserForm1.Label1.Width = 0
   UserForm1.Label_���� = "����������"
  'DoEvents
  �����_������_�_�����������_������_���������_��������� = True
  UserForm1.Show
  �����_������_�_�����������_������_���������_��������� = False
  'UserForm1.Hide
End Sub

Sub ��������_��������_����_�������()
  'UserForm1.Repaint
   UserForm1.Label1.Width = 0
   UserForm1.Label_���� = "����������"
  'DoEvents
  �����_��������_����_������� = True
  UserForm1.Show
  �����_��������_����_������� = False
  'UserForm1.Hide
End Sub


Sub �_����������()
Dim r As Range, x, cl As New Collection, s$
  Set r = ActiveDocument.Range
  On Error Resume Next
  With r.Find
    .Text = "<[�-ߨ]{2;}>" '���� [0-9.-]{2;} <[�-ߨ]{2;}>
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


Sub ��������_����_�������()
t = Timer
�����_������_�_����� = ActiveDocument.ComputeStatistics(wdStatisticPages)

For i = 1 To ActiveDocument.Tables.Count Step 1 '������ �� ��������
  For j = 1 To ActiveDocument.Tables(i).Rows.Count  '������ �� �������
     If ActiveDocument.Tables(i).Columns.Count = 18 Then
        ActiveDocument.Tables(i).Cell(j, 2).Select
        
       DoEvents
       �����_��������_��� = Selection.Information(wdActiveEndPageNumber) '����� �������� ���       DoEvents
       UserForm1.Label1.Width = CInt((300 * �����_��������_���) / �����_������_�_�����)
       UserForm1.Label_���� = "���� " & �����_��������_��� & " �� " & �����_������_�_�����
       'If (Timer - t) > 60 Then
       '   UserForm1.Label2 = "����� " & Round((Timer - t) / 60, 1) & " ���" ' �����  � ���
       'Else
       '   UserForm1.Label2 = "����� " & Round(Timer - t, 1) & " ���" ' �����  � ���
       'End If
       UserForm1.Label2 = "�����: " & TimeSerial(0, 0, Timer - t)
       
       UserForm1.Repaint
       
        Call ����_�����
     End If
  Next j
Next i
End Sub


Sub ����������_�_������()
Dim objExcApp As Object
'�1 = Selection.Rows.First.Index  '����� �������� ������ � �������
'�1 = Selection.Columns.First.Index  '����� ��������� ������� � �������
 �1 = Selection.Information(wdEndOfRangeRowNumber) '����� ������ ������ ������� � ������� ������ ������
 �1 = Selection.Information(wdEndOfRangeColumnNumber) '����� ������� ������ ������� � ������� ������ ������


'����������� ������ �������� �������
Set tblSel = Selection.Tables(1)
ingStart = tblSel.Range.Start
For i = 1 To ActiveDocument.Tables.Count Step 1
 If ActiveDocument.Tables(i).Range.Start = ingStart Then
    ingTbIndex = i
    Exit For
 End If
 Next i
 
�����_�� = ActiveDocument.Tables(ingTbIndex).Cell(�1, �1).Range.Text    '������� �������� � ������ � �������
�����_�� = Replace(�����_��, Chr(13) & "", "")
�����_�� = Replace(�����_��, ",", ".")

'������ � �������

On Error Resume Next '���������� ������
Set objExcApp = GetObject(, "Excel.Application")
  '���� ������ �� ������
  If objExcApp Is Nothing Then
     On Error GoTo 0  '����� �� ���������� ������
     Set objExcApp = CreateObject("Excel.Application")
'     objExcApp.Workbooks.Add '������� ����� ������
     'objExcApp.Visible = True '������� ������ �������
     '������� ������ � ������ ������� ����
     On Error Resume Next
     ��������� = objExcApp.Application.Evaluate(�����_��)
     ��������� = Replace(���������, ".", ",")
     ActiveDocument.Tables(ingTbIndex).Cell(�1, �1).Range = ���������
     On Error GoTo 0  '����� �� ���������� ������
     objExcApp.Quit '������� ��������� ������
     GoTo �_�����
  '���� ������ ������
  Else
     On Error Resume Next
     ��������� = objExcApp.Application.Evaluate(�����_��)
     ��������� = Replace(���������, ".", ",")
     ActiveDocument.Tables(ingTbIndex).Cell(�1, �1).Range = ���������
     On Error GoTo 0  '����� �� ���������� ������
'     Set objExcDoc = objExcApp.Workbooks.Application
'     objExcDoc.Sheets.Add After:=objExcDoc.Sheets(objExcDoc.Sheets.Count) '������� ���� � ����� ������
     'objExcApp.Visible = True '������� ������ �������
  End If
  
�_�����:

Call ����_�����

End Sub



Sub ����_��������_���()
Dim objExcApp As Object
'Dim ie As Object


Dim Shell_Object

If x_���� = ������������.Count Or x_���� = 0 Then
'Set ie = CreateObject("InternetExplorer.Application")
Set objshell = CreateObject("Wscript.shell")
Set ie = New InternetExplorerMedium
End If

'������ = "���� 19903-"
If ������_����� = True Then
  ����_������� = �������_����
Else
  ����_������� = Selection.Text
End If

����_������� = Replace(����_�������, Chr(13) & "", "")

������ = Split(����_�������, "-")
�����_���� = UBound(������, 1)

������ = "" '���� ��� ����
If �����_���� <> 0 Then
  For i = 0 To �����_���� - 1
     ������ = ������ + ������(i) + "-"
  Next i
Else
������ = ����_�������
End If


'��������
If x_���� = ������������.Count Or x_���� = 0 Then
  ie.Silent = True
  ie.Visible = True
  ie.Navigate "http://i1:8085/idoc/client/jsp/main.jsp?trail=~C~1~A~2~S~child_oks~C~3~A~2~S~child_oks~C~4~A~2~S~child_oks~C~403586~A~2#~C~1~A~2~S~child_oks~C~3~A~2~V~2~C~3~A~2"
  
  '
  Do While ie.Busy = True Or ie.ReadyState <> 4: DoEvents: Loop
End If

ie.Document.getElementById("findPatt").Value = ������

'Debug.Print ie.Document.getElementsByTagName("td")(2).getElementsByTagName("input")(0).getAttribute("alt")
'Debug.Print ie.Document.getElementsByTagName("input")(3).getAttribute("alt")

��������� = Len(ie.Document.body.innerText)


'Debug.Print ���������

'���� ���� ���������� �������� ������ �� 26155

'If x_���� = ������������.Count Or x_���� = 0 Then
'  ���������� = False
'  Do While ���������� <> True
'   �������� = Len(ie.Document.body.innerText)
'   If �������� = 26155 Or �������� = 26336 Then ���������� = True
'   DoEvents
'  Loop
'Else
  '���� ��������� �����
  ����� = Timer + 3 '����� ��� �������
  Do While Timer < �����
   DoEvents
   'Debug.Print Len(ie.Document.body.innerText)
  Loop
'End If



'���� ��������� �����
����� = Timer + 1 '����� ��� �������
Do While Timer < �����
 DoEvents
 'Debug.Print Len(ie.Document.body.innerText)
Loop

����������� = Len(ie.Document.body.innerText)
'Debug.Print �����������

'��������� = Len(ie.Document.body.innerText)
'Debug.Print ���������



ie.Document.getElementsByTagName("input")(3).Click

'�������� = Len(ie.Document.body.innerText)
'Debug.Print ��������


'���� ���� ���������� �������� ������ �� 4005
���������� = False
Do While ���������� <> True
 �������� = Len(ie.Document.body.innerText)
 If �������� <> 4005 Then ���������� = True
 DoEvents
Loop

'���� ��������� �����
����� = Timer + 2 '����� ��� �������
Do While Timer < �����
 DoEvents
 'Debug.Print Len(ie.Document.body.innerText)
Loop

' ����������� = Len(ie.Document.body.innerText)
' Debug.Print �����������
 

�������������� = ie.Document.body.innerText
If InStr(1, ��������������, "������ �� �������") <> 0 Then
  ���������_���� = ""
  Exit Sub
End If
'��������������HTML = ie.Document.body.innerHTML
'Debug.Print ��������������
'Debug.Print ��������������HTML


'����������� ����� ������ ���

�����_������� = Len(������)
������_��� = "0"

����� = 0
��� = "0"

�����_����_���:
  
  If ��� = "" Then ��� = "0"
  If ������_��� = "" Then ������_��� = "0"
  If CSng(Replace(������_���, "-", "")) < CSng(Replace(���, "-", "")) Then ������_��� = ���
  'On Error GoTo 0  '����� �� ���������� ������

  '������_����� = �����
  ��� = ""
  ����� = 0
  �_����_���� = InStr(1 + �_����_����, ��������������, ������)
  If �_����_���� = 0 Then GoTo ���������_�����_����
  �_�������_���� = �_����_���� + �����_�������
  For i = �_�������_���� To Len(��������������)
    DoEvents
    ������ = Mid(��������������, �_�������_���� + �����, 1)
    If ������ Like "[0-9-]" Then '����� ����� � ����
       ��� = ��� + ������
       ����� = ����� + 1
    Else '���� �� �����
      GoTo �����_����_���
      'Exit For
    End If
    DoEvents
  Next i

���������_�����_����:


���������_���� = ������ + ������_���
If ������_����� = False Then
  Selection.Text = ���������
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


'objshell.SendKeys "���� 2.729-68"
'objshell.SendKeys "{Enter}"
'objshell.SendKeys "~"
If ������_����� = False Then
  ie.Quit
  Set objshell = Nothing
  Set ie = Nothing
  Set ������������ = Nothing
End If


End Sub


Sub ��������_�����_������()
Dim r As Range, x, s$
t = Timer

Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=1, Name:="" '������� �� ��� �� ������
  
ReDim ������_�����(2) '������ ������ �������
'������� ����� ������ (���� � ���������)
������_�����(0) = "����[0-9.-]{2;}"
������_�����(1) = "<[�]�� [0-9.-]{2;}"
������_�����(2) = "<[�]� [0-9.-]{2;}"
'���� ����� �����
For i = 0 To UBound(������_�����, 1)
Dim cl As Collection
  Set r = ActiveDocument.Range
  Set cl = New Collection
  On Error Resume Next
  With r.Find
    .Text = ������_�����(i) '���� [0-9.-]{2;} <[�-ߨ]{2;}>
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

' �������� ����� ����� �� ����������
For x_���� = cl.Count To 1 Step -1
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = cl.Item(x_����)
        If i = 0 Then
          .Replacement.Text = Replace(cl.Item(x_����), "����", "���� ")
          .Replacement.Text = Replace(.Replacement.Text, "����" & Chr(160), "���� ")
        End If
        If i = 1 Then
          .Replacement.Text = Replace(cl.Item(x_����), "��� ", "���")
          .Replacement.Text = Replace(.Replacement.Text, "���" & Chr(160), "���")
        End If
        If i = 2 Then
          .Replacement.Text = Replace(cl.Item(x_����), "�� ", "��")
          .Replacement.Text = Replace(.Replacement.Text, "��" & Chr(160), "��")
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
Next x_����
Next i

Set cl = Nothing
  
  
'���� � ��������� ���� � ������ �� �������
  ����� = 9
  ReDim ������_������(�����) '������ ������ �������
  ����� = ����� - �����
  '������� ������
  ������_������(�����) = "���� [0-9.-]{2;}":         ����� = ����� + 1
  ������_������(�����) = "���� � ��� [0-9.-]{2;}":   ����� = ����� + 1
  ������_������(�����) = "���� [�-ߨ] [0-9.-]{2;}":  ����� = ����� + 1
  ������_������(�����) = "���� �� [0-9.-]{2;}":      ����� = ����� + 1
  ������_������(�����) = "���[0-9.-]{2;}":           ����� = ����� + 1  '���� [0-9.�]{2;}-[0-9]{2;}
  ������_������(�����) = "����5�[0-9.-]{2;}":        ����� = ����� + 1  '���� [0-9.�]{2;}-[0-9]{2;}
  ������_������(�����) = "����5[0-9.-]{2;}":         ����� = ����� + 1  '���� [0-9.�]{2;}-[0-9]{2;}
  ������_������(�����) = "���5�[0-9.-]{2;}":         ����� = ����� + 1  '���� [0-9.�]{2;}-[0-9]{2;}
  ������_������(�����) = "���5[0-9.-]{2;}":          ����� = ����� + 1  '���� [0-9.�]{2;}-[0-9]{2;}
  ������_������(�����) = "��5[0-9.-]{2;}":           ����� = ����� + 1  '���� [0-9.�]{2;}-[0-9]{2;}

  Set r = ActiveDocument.Range
  Set ������������ = New Collection
  On Error Resume Next
  
For i = 0 To UBound(������_������, 1)
  With r.Find
    .Text = ������_������(i) '���� [0-9.-]{2;} <[�-ߨ]{2;}>
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
      For x = 1 To ������������.Count
        DoEvents
        If s < ������������(x) Then ������������.Add s, s, Before:=x: GoTo 1
      Next
      ������������.Add s, s
1     .Wrap = wdFindStop 'wdFindContinue 'wdFindStop
   Wend
  End With
  
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
Next i

On Error GoTo 0  '����� �� ���������� ������


' ����� � �����


For x_���� = ������������.Count To 1 Step -1
  
       Application.ScreenUpdating = True '�������� ���������� ���������
       DoEvents
       Application.ScreenUpdating = False '�������� ���������� ���������
       UserForm1.Label1.Width = CInt((300 * (������������.Count - x_���� + 1)) / ������������.Count)
       UserForm1.Label_���� = "����������� ��������: " & ������������.Item(x_����) & "  (" & ������������.Count - x_���� + 1 & " �� " & ������������.Count & ")"
       UserForm1.Label2 = "�����: " & TimeSerial(0, 0, Timer - t)
       
       UserForm1.Repaint
  
  
  �������_���� = ������������.Item(x_����)
  ������_����� = True
  Call ����_��������_���
  ������_����� = False
  ���_����� = ���������_����
  If ���������_���� = "" Then
    Options.DefaultHighlightColorIndex = wdRed
    ���������_���� = �������_����
  Else
    Options.DefaultHighlightColorIndex = wdBrightGreen
  End If
  '������������.Item(x_����) = ���������_����
    
    Selection.Find.ClearFormatting
    Selection.Find.Highlight = False '�������� ����� �� �������������
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Highlight = True '������ ����� �������� ������
    'Selection.Find.Font.Color = 10498160
    With Selection.Find
        .Text = �������_����
        .Replacement.Text = ���������_����
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



Sub ��������_��������_�����_������()
  Application.ScreenUpdating = False '��������� ���������� ���������
  'UserForm1.Repaint
   UserForm1.Label1.Width = 0
   UserForm1.Label_���� = "����������"
  'DoEvents
  �����_��������_�����_������ = True
  UserForm1.Show
  �����_��������_�����_������ = False
  'UserForm1.Hide
  On Error Resume Next
  ie.Quit
  Set objshell = Nothing
  Set ie = Nothing
  Set ������������ = Nothing
  On Error GoTo 0  '����� �� ���������� ������
  
  Application.ScreenUpdating = True '�������� ���������� ���������
End Sub






Sub ���������_����_�_�������_�_����������()
t = Timer
    
'If �����������_������_��_�������_��� = True Then
' GoTo �����������_������_��_�������
'End If
    
 ' If �����������_����_������ = False Then      '�� �������� ������ �����_������_���_��_����
 '    ��������� = MsgBox("Yes-""���."", No-""���."", Cancel-����� ", vbYesNoCancel + vbQuestion + vbDefaultButton1, "������ �� ������ �������")
 '    If ��������� = vbCancel Then Exit Sub
 '    If ��������� = vbYes Then ���� = "���."
 '    If ��������� = vbNo Then ���� = "���."
 ' End If
  
  ���� = "���."
  
'�����������_������_��_�������:
    
 '   If �����������_����_������ = False Then
       Application.ScreenUpdating = False '��������� ���������� ���������
 '   End If
 
 �����_������_�_��������� = ActiveDocument.ComputeStatistics(wdStatisticPages)
    
 If Selection.PageSetup.PageWidth = CSng(Format(CentimetersToPoints(21), "0.0")) And _
      Selection.PageSetup.PageHeight = CSng(Format(CentimetersToPoints(29.7), "0.0")) And _
      Selection.PageSetup.VerticalAlignment = 0 _
 Then
      ������_����� = "�4"
 Else
      ������_����� = "�3"
 End If
    
    
If ������_����� = "�4" Then
   �������_������� = ActiveWindow.ActivePane.View.Zoom.Percentage
   ActiveWindow.ActivePane.View.Zoom.Percentage = 150
End If

   �����_��������_��� = Selection.Information(wdActiveEndPageNumber) '����� �������� ���

      
      Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=�����_��������_���, Name:="" '������� �� ��� �� ������
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
      WordBasic.ViewFooterOnly '������� �� ������ ����������
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
      WordBasic.ViewFooterOnly '������� �� ������ ����������
      �����_�����_�������_���� = Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text
      �����_�����_�������_���� = Replace(�����_�����_�������_����, Chr(13) & "", "")
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
      
      Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=�����_��������_��� + 1, Name:="" '������� �� ��� �� ������
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
      WordBasic.ViewFooterOnly '������� �� ������ ����������
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
      WordBasic.ViewFooterOnly '������� �� ������ ����������
      �����_�����_�������_���� = Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text
      �����_�����_�������_���� = Replace(�����_�����_�������_����, Chr(13) & "", "")
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������

      If ������_����� = "�4" Then
         �����_������_����� = 3
      Else
         �����_������_����� = 4
      End If

'���� False ��������� ����� ������� ���
 'If �����������_����_������ = False Then     '�� �������� ������ �����_������_���_��_����
   Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=�����_��������_���, Name:="" '������� �� ��� �� ������
   With ActiveDocument.Range.Tables(ActiveDocument.Range.Tables.Count)
      Erase ���� '��������� ������
      ReDim ����(.Rows.Count * 2, .Rows.Count * 2) '������ ������ �������
      '��������� ���������
      For i = �����_������_����� To .Rows.Count   ' ��������� ��� �� ��������� ���� ��� ������� ��� (�1, �2 ...)
         If .Cell(i, 1).Range.Text <> Chr(13) & "" Then
            ����(a1, 1) = Left(.Cell(i, 1).Range.Text, Len(.Cell(i, 1).Range.Text) - 2)
            '���� = "���."
            ����(a1, 4) = Left(.Cell(i, 7).Range.Text, Len(.Cell(i, 7).Range.Text) - 2)
            ����(a1, 5) = Left(.Cell(i, 9).Range.Text, Len(.Cell(i, 9).Range.Text) - 2)
            ����(a1, 6) = Left(.Cell(i, 10).Range.Text, Len(.Cell(i, 10).Range.Text) - 2)
         End If
      Next i
   End With
' End If
 

'+++��������� ������ �������� ���� ��� ��� � ������� ����� "��� � ���������� ������� -  �������� �� �������"

    
   ' Selection.MoveUp Unit:=wdLine, Count:=1, Extend:=wdExtend

    
  '  Selection.Find.ClearFormatting
  '  f = False
  '  With Selection.Find
  '      .Text = "^b"
  '      .Replacement.Text = ""
  '      .Forward = True
  '      .Wrap = wdFindStop 'wdFindStop �� ����������� ���� �� �����
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
  '    Selection.InsertBreak Type:=wdSectionBreakContinuous  '������ �� ������� �������� �� ��������� ���
 '  Else
  '    Selection.MoveRight Unit:=wdCharacter, Count:=1
 '   End If
    
'+++��������� ������ �������� ���� ��� ��� � ������� ����� "��� � ���������� ������� -  �������� ����� �������"
      '������� ����� ��������� ���
    Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=�����_��������_��� + 1, Name:="" '������� �� ��� �� ������
    Selection.MoveUp Unit:=wdLine, Count:=1, Extend:=wdExtend
    Selection.Find.ClearFormatting
    f = False
    With Selection.Find
        .Text = "^b"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop 'wdFindStop �� ����������� ���� �� �����
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
      If ������_����� = "�4" Then
         Selection.HomeKey Unit:=wdLine '������� � ������ ������
         Selection.InsertBreak Type:=wdSectionBreakNextPage    '������ �� ��������� ��������
      Else
         Selection.MoveUp Unit:=wdLine, Count:=1
         Selection.InsertBreak Type:=wdSectionBreakContinuous  '������ �� ������� �������� �� ��������� ���
      End If
    Else
     Selection.MoveRight Unit:=wdCharacter, Count:=1
    End If
      Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=�����_��������_��� + 1, Name:="" '������� �� ��� �� ������
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
      WordBasic.ViewFooterOnly '������� �� ������ ����������
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
      WordBasic.ViewFooterOnly '������� �� ������ ����������
      'WordBasic.ViewFooterOnly '������� �� ������ ����������
      Selection.HeaderFooter.LinkToPrevious = False ' ����� "��� � ��������� �������"
      �����_�����_�������_��� = Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text
      �����_�����_�������_��� = Replace(�����_�����_�������_���, "", "")
      �����_�����_�������_��� = Replace(�����_�����_�������_���, Chr(13), "")
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������

      Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=�����_��������_���, Name:="" '������� �� ��� �� ������
      WordBasic.ViewFooterOnly ' ������� ������ ����������
      
      �����_�����_�������_��� = Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text
      �����_�����_�������_��� = Replace(�����_�����_�������_���, "", "")
      �����_�����_�������_��� = Replace(�����_�����_�������_���, Chr(13), "")
      �����_�����_�������_���_���� = Selection.HeaderFooter.PageNumbers.StartingNumber
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
      
     '������� ����� ���������� ���
 '   If �����_��������_��� <> �����_��������_��� Then
  '    Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=�����_��������_��� - 1, Name:="" '������� �� ��� �� ������
  '    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
  '    WordBasic.ViewFooterOnly '������� �� ������ ����������
  '    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
  '    WordBasic.ViewFooterOnly '������� �� ������ ����������
  '    �����_�����_�����_��� = Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text
  '    �����_�����_�����_��� = Replace(�����_�����_�����_���, "", "")
  '    �����_�����_�����_��� = Replace(�����_�����_�����_���, Chr(13), "")
 '     ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
 '     Selection.GoTo What:=wdGoToPage, Which:=wdGoToNext, Name:="+1"
 '   End If
      
      ' �����_�����_�����_���
      ' �����_�����_�������_���
      ' �����_�����_�������_���
      
      ' �����_�����_�����_����
      ' �����_�����_�������_����
      ' �����_�����_�������_����

    '  If �����_�����_�����_��� <> �����_�����_�����_���� Then
    '        Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=�����_��������_��� - 1, Name:="" '������� �� ��� �� ������
    '       WordBasic.ViewFooterOnly ' ������� ������ ����������
    '         With Selection.HeaderFooter.PageNumbers  ' ���������� �������
    '           .NumberStyle = wdPageNumberStyleArabic
    '           .HeadingLevelForChapter = 0
    '           .IncludeChapterNumber = False
    '           .ChapterPageSeparator = wdSeparatorHyphen
    '           .RestartNumberingAtSection = False
    '           .StartingNumber = 0
   '         End With
   '         If Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text, Chr(13) & "", "") <> �����_�����_�����_���� Then
   '            With Selection.HeaderFooter.PageNumbers  ' ������ ����� ��� �������
   '              .NumberStyle = wdPageNumberStyleArabic
   '              .HeadingLevelForChapter = 0
    '             .IncludeChapterNumber = False
   '              .ChapterPageSeparator = wdSeparatorHyphen
   '              .RestartNumberingAtSection = True
   '              .StartingNumber = �����_�����_�����_����
   '           End With
   '        End If
   '        ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
   '   End If
      
      If �����_�����_�������_��� <> �����_�����_�������_���� Then
            Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=�����_��������_���, Name:="" '������� �� ��� �� ������
           WordBasic.ViewFooterOnly ' ������� ������ ����������
             With Selection.HeaderFooter.PageNumbers  ' ���������� �������
               .NumberStyle = wdPageNumberStyleArabic
               .HeadingLevelForChapter = 0
               .IncludeChapterNumber = False
               .ChapterPageSeparator = wdSeparatorHyphen
               .RestartNumberingAtSection = False
               .StartingNumber = 0
            End With
            If Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text, Chr(13) & "", "") <> �����_�����_�������_���� Then
               With Selection.HeaderFooter.PageNumbers  ' ������ ����� ��� �������
                 .NumberStyle = wdPageNumberStyleArabic
                 .HeadingLevelForChapter = 0
                 .IncludeChapterNumber = False
                 .ChapterPageSeparator = wdSeparatorHyphen
                 .RestartNumberingAtSection = True
                 .StartingNumber = ������_�������(�����_��������_��� - 1)
              End With
           End If
           If Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text, Chr(13) & "", "") <> �����_�����_�������_���� Then
               Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text = �����_�����_�������_����
           End If
           ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
      End If
      
      
      'If �����_�����_�������_��� <> �����_�����_�������_���� Then
      ' ������ �������� ��� �� "�����_�����_�������" ���
           Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=�����_��������_��� + 1, Name:="" '������� �� ��� �� ������
           WordBasic.ViewFooterOnly ' ������� ������ ����������
               With Selection.HeaderFooter.PageNumbers  ' ������ ����� ��� �������
                 .NumberStyle = wdPageNumberStyleArabic
                 .HeadingLevelForChapter = 0
                 .IncludeChapterNumber = False
                 .ChapterPageSeparator = wdSeparatorHyphen
                 .RestartNumberingAtSection = True
                 .StartingNumber = �����_�����_�������_����
              End With
        '   End If
           ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
      'End If

      
'++++++++'������� ����� �������� (�������� ����� ���� ��� �� ������ �� ��������� ���)

If ������_����� = "�4" Then
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.EndKey Unit:=wdLine '������ � ����� ������
    Selection.InsertBreak Type:=wdSectionBreakNextPage    '������ �� ��������� ��������
    Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=�����_��������_��� + 1, Name:="" '������� �� ��� �� ������
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
    WordBasic.ViewFooterOnly '������� �� ������ ����������
    Selection.HeaderFooter.LinkToPrevious = False ' ����� "��� � ��������� �������"
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
    WordBasic.ViewFooterOnly '������� �� ������ ����������
Else
    'ActiveWindow.ActivePane.View.NextHeaderFooter '��������� ������
    'ActiveWindow.ActivePane.View.PreviousHeaderFooter '���������� ������
    
    '���������� ���������� ����� �� ��������
    �����_�������_��� = �����_��������_���
    ���������_�����_��_��� = 0
     Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=�����_��������_���, Name:="" '������� �� ��� �� ������
     Do While �����_��������_��� + 1 <> �����_�������_���
          Selection.MoveDown Unit:=wdLine, Count:=1 '�� ����� � ����
          ���������_�����_��_��� = ���������_�����_��_��� + 1
          �����_�������_��� = Selection.Information(wdActiveEndPageNumber) '����� ������� ���
          DoEvents
     Loop
     ���������_�����_��_��� = ���������_�����_��_��� - 1
     

     
     '��������� ������ � ����� ��������� ��������
     Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=�����_��������_��� + 1, Name:="" '������� �� ��� �� ������
     Selection.MoveUp Unit:=wdLine, Count:=1 '�� ����� � ����
     Selection.MoveUp Unit:=wdLine, Count:=1 '�� ����� � ����
     
     '�����_������_�_����� = ActiveDocument.ComputeStatistics(wdStatisticPages)
     'If �����_��������_��� = �����_������_�_����� - 1 Then
     '   Selection.MoveUp Unit:=wdLine, Count:=1 '�� ����� � ����
     'End If
     
    '������ ����� � ��������� ������ ���� �� ��������� �����
On Error Resume Next '���������� ������
    Selection.Cells(1).Row.Select
    If Err.Number = 5991 Then '���� ������ � �������� ����������
       '����������� ������ �������� �������
       Set tblSel = Selection.Tables(1)
       ingStart = tblSel.Range.Start
       For i = 1 To ActiveDocument.Tables.Count Step 1
          If ActiveDocument.Tables(i).Range.Start = ingStart Then
            ingTbIndex = i
            Exit For
          End If
       Next i
     '�������� ��� ������ �������
       For i = 1 To ActiveDocument.Tables(ingTbIndex).Columns.Count
         Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
       Next i
    End If
On Error GoTo 0  '����� �� ���������� ������
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

   '������ ������� ����� ����� �������
    Selection.Borders(wdBorderRight).LineStyle = wdLineStyleNone


     Selection.InsertRowsBelow ���������_�����_��_��� '�������� ������ � �������
     '������ ����� ����� ����� �������
     Selection.Borders(wdBorderTop).LineStyle = wdLineStyleNone
     
     '�������� ���������� �� ������ ������� �����
     Do While �����_��������_��� + 1 = �����_�������_���
          Selection.InsertRowsBelow 1 '�������� ������ � �������
          '���������_�����_��_��� = ���������_�����_��_��� + 1
          �����_�������_��� = Selection.Information(wdActiveEndPageNumber) '����� ������� ���
          DoEvents
     Loop
     Selection.Rows.Delete


    Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=�����_��������_��� + 1, Name:="" '������� �� ��� �� ������
    Selection.InsertBreak Type:=wdSectionBreakContinuous  '������ �� ������� �������� �� ��������� ���

     
   '����� "��� � ��������� �������"
   WordBasic.ViewFooterOnly ' ������� ������ ����������
    If Selection.HeaderFooter.LinkToPrevious = True Then ' ����� "��� � ��������� �������"
       Selection.HeaderFooter.LinkToPrevious = False ' ����� "��� � ��������� �������"
       ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� ����������� ���� �� ����� �� ����������� � ����� �� ����� ������ � ��������� ����� �� ���� ����������
       WordBasic.ViewFooterOnly ' ������� ������ ����������
       ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    End If

    '������ ����� � ��������� ������ ���� �� ��������� �����
     Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=�����_��������_��� + 1, Name:="" '������� �� ��� �� ������
     Selection.MoveUp Unit:=wdLine, Count:=1 '�� ����� � ����
     Selection.MoveUp Unit:=wdLine, Count:=1 '�� ����� � ����
     On Error Resume Next '���������� ������
         Selection.Cells(1).Row.Select
         If Err.Number = 5991 Then '���� ������ � �������� ����������
            '����������� ������ �������� �������
            Set tblSel = Selection.Tables(1)
            ingStart = tblSel.Range.Start
            For i = 1 To ActiveDocument.Tables.Count Step 1
               If ActiveDocument.Tables(i).Range.Start = ingStart Then
                 ingTbIndex = i
                 Exit For
               End If
            Next i
          '�������� ��� ������ �������
            For i = 1 To ActiveDocument.Tables(ingTbIndex).Columns.Count
              Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
            Next i
         End If
     On Error GoTo 0  '����� �� ���������� ������
     With Selection.Borders(wdBorderBottom)
         .LineStyle = Options.DefaultBorderLineStyle
         .LineWidth = Options.DefaultBorderLineWidth
         .Color = Options.DefaultBorderColor
     End With

End If '������_����� = "�4"

    Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=�����_��������_��� + 1, Name:="" '������� �� ��� �� ������
    WordBasic.ViewFooterOnly ' ������� ������ ����������
    Selection.HeaderFooter.LinkToPrevious = False ' ����� "��� � ��������� �������"
  
'ReDim ����(1, 6) '������ ������ �������
'����(a1, 1) = "�1"
'���� = "���."
'����(a1, 4) = "22220.43.___"
'����(a1, 5) = "�����������"
'����(a1, 6) = "15.16.18"


'���� ����� �������� ������ ���
������_�����_������� = 0.1
�����_������� = �����_�����_�������_���
If InStr(1, �����_�����_�������_���, ".") > 0 Then
   �����_������� = CSng(Split(�����_�����_�������_���, ".")(1))
   �����_������� = CSng(Split(�����_�����_�������_���, ".")(0))
   ������_�����_������� = 10 ^ (-1 * CSng(Len(�����_�������)))
End If
'������ ����� ���
'Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Select
'Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:="PAGE  \* Arabic ", PreserveFormatting:=True

'��������� ����� ��������
  Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text = CStr(�����_�������) + "."
  Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Select
  Selection.EndKey Unit:=wdLine '������ � ����� ������
  Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:="PAGE  \* Arabic ", PreserveFormatting:=True  '���������� ����� ����� ��������



If InStr(1, �����_�����_�������_���, ".") = 0 Then  '���� �� ��������� ��� ������ ����� ����� ����� � ������ ���
               With Selection.HeaderFooter.PageNumbers  ' ������ ����� ��� �������
                 .NumberStyle = wdPageNumberStyleArabic
                 .HeadingLevelForChapter = 0
                 .IncludeChapterNumber = False
                 .ChapterPageSeparator = wdSeparatorHyphen
                 .RestartNumberingAtSection = True
                 .StartingNumber = 1
              End With
            
            If Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text, Chr(13) & "", "") = Replace(�����_�����_�������_��� + ������_�����_�������, ",", ".") Then
               GoTo �����_�����
            End If

Else '���� ����
             With Selection.HeaderFooter.PageNumbers  ' ���������� �������
               .NumberStyle = wdPageNumberStyleArabic
               .HeadingLevelForChapter = 0
               .IncludeChapterNumber = False
               .ChapterPageSeparator = wdSeparatorHyphen
               .RestartNumberingAtSection = False
               .StartingNumber = 0
            End With
            
            If Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text, Chr(13) & "", "") = Replace(�����_�����_�������_��� + ������_�����_�������, ",", ".") Then
               GoTo �����_�����
            End If
            
               With Selection.HeaderFooter.PageNumbers  ' ������ ����� ��� �������
                 .NumberStyle = wdPageNumberStyleArabic
                 .HeadingLevelForChapter = 0
                 .IncludeChapterNumber = False
                 .ChapterPageSeparator = wdSeparatorHyphen
                 .RestartNumberingAtSection = True
                 .StartingNumber = �����_������� + 1
              End With
              
            If Replace(Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text, Chr(13) & "", "") = Replace(�����_�����_�������_��� + ������_�����_�������, ",", ".") Then
               GoTo �����_�����
            End If
            
End If

Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text = Replace(�����_�����_�������_��� + ������_�����_�������, ",", ".")

�����_�����:

'If Replace(�����_������� + 1, ",", ".") = �����_�����_�������_���_���� Then
'  Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Text = �����_������� + "."
'  Selection.EndKey Unit:=wdLine '������ � ����� ������
'  Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:="PAGE  \* Arabic ", PreserveFormatting:=True
  'Selection.HomeKey Unit:=wdLine '������ � ������ ������
  'Selection.TypeText Text:=�����_�������
'Else
'  Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text = Replace(�����_�����_�������_��� + ������_�����_�������, ",", ".")
'End If
'�����_�����_�������_���






'Selection.HeaderFooter.PageNumbers.StartingNumber '����� ��� ���������� ����� (��������� �������/������ ������� �������....)



'��������� ���������� �������
  If ���������_����� = False Then
   '   Selection.Find.Execute '�������� �������� �����  'Dim ����()
      'Selection.HeaderFooter.Range.Tables(1).Cell(2, 7).Range.Text
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 1).Range.Text = ����(a1, 1) '������� � �����������
      'Selection.Tables(1).Cell(2, 1).Range.Text = "�1" '������� � �����������
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 2).Range.Text = ���� '������� � �����������
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 3).Range.Text = ����(a1, 4) '������� � �����������
      'Selection.Tables(1).Cell(2, 3).Range.Text = "22220.43.___" '������� � �����������
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 3).Select
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 3).FitText = True
      'Selection.HeaderFooter.Range.Cells(1).FitText = True  '������ �� ������ ������
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 4).Range.Text = ����(a1, 5) '������� � �����������
      'Selection.Tables(1).Cell(2, 4).Range.Text = "�����������" '������� � �����������
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 4).Select
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 4).FitText = True  '������ �� ������ ������
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 5).Range.Text = ����(a1, 6) '������� � �����������
      'Selection.Tables(1).Cell(2, 5).Range.Text = "15.16.18" '������� � �����������
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 5).Select
      Selection.HeaderFooter.Range.Tables(1).Cell(2, 5).FitText = True
  End If
  

      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' ����� �� �����������
  
    
    If (Timer - t) > 60 Then
      Debug.Print (Timer - t) / 60 & " ���" ' �����  � ���
      Else
      Debug.Print Timer - t & " ���" ' �����  � ���
    End If
  
    
�_������:

If ������_����� = "�4" Then
    ActiveWindow.ActivePane.View.Zoom.Percentage = �������_�������
End If

Application.ScreenUpdating = True '�������� ���������� ���������
    

End Sub



Sub ������_����()
Attribute ������_����.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.����"


'����������� ������ �������� �������
Set tblSel = Selection.Tables(1)
ingStart = tblSel.Range.Start
For i = 1 To ActiveDocument.Tables.Count Step 1
 If ActiveDocument.Tables(i).Range.Start = ingStart Then
    ����_ingTbIndex = i
    Exit For
 End If
Next i
 '����������� ��������� �������� ������
 ����_�1 = Selection.Information(wdEndOfRangeRowNumber) '����� ������ ������ ������� � ������� ������ ������
 ����_�1 = Selection.Information(wdEndOfRangeColumnNumber) '����� ������� ������ ������� � ������� ������ ������
   
   Call ������_����_�������

End Sub

Sub ������_����_�������()
    '��������� ������� �������
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyReturn), _
        KeyCategory:=wdKeyCategoryCommand, Command:="������_����_���������_�����"
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyEsc), _
        KeyCategory:=wdKeyCategoryCommand, Command:="������_����_���������_�����_������"
        
        MsgBox ("�������� ����� ������� � ������� Enter." & vbNewLine & vbNewLine & "��� ������ ������� �� ����� Esc")
        
End Sub

Sub ������_����_���������_�����_������()
'����� ������� �������
CustomizationContext = NormalTemplate
FindKey(BuildKeyCode(Arg1:=wdKeyReturn)).Clear
CustomizationContext = NormalTemplate
FindKey(BuildKeyCode(Arg1:=wdKeyEsc)).Clear
Err.Clear
End Sub


Sub ������_����_���������_�����()
Attribute ������_����_���������_�����.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.����"

On Error Resume Next '���������� ������
    Selection.Copy
      ActiveDocument.Tables(����_ingTbIndex).Cell(����_�1 - 1, ����_�1).Select ' ������������ �� ���� ������� ���� ��������
      Selection.HomeKey Unit:=wdLine '������� � ������ ������ (������)
      'Selection.SplitTable ' ������� �������
      'Selection.InsertBreak Type:=wdSectionBreakNextPage  '������ �� ��������� �������� �� ��������� ���
      Selection.InsertBreak Type:=wdPageBreak '������� ������ ��������
      ActiveDocument.Tables(����_ingTbIndex + 1).Cell(1, 1).Select
      Selection.HomeKey Unit:=wdLine '������� � ������ ������
    Selection.Paste
    ActiveDocument.Tables(����_ingTbIndex + 1).Cell(1, 1).Select
    Selection.InsertRowsAbove 1 '�������� ������� ������
    Selection.Cells.Merge ' ���������� ������
    ' �����
    Selection.Borders(wdBorderTop).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderLeft).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderRight).LineStyle = wdLineStyleNone
    
    Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft '������������ �� ������ ����
    Selection.TypeText Text:="����������� ������� " '������� �����
    
'����� ������� �������
CustomizationContext = NormalTemplate
FindKey(BuildKeyCode(Arg1:=wdKeyReturn)).Clear
CustomizationContext = NormalTemplate
FindKey(BuildKeyCode(Arg1:=wdKeyEsc)).Clear
Err.Clear

On Error GoTo 0  '����� �� ���������� ������
End Sub

Sub �����_����������()
    '������ ������ �������� ������
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Highlight = False
    Selection.Find.Replacement.Text = ""
    
    '���� ������� � ���������������� ���� � ��������
    Options.DefaultHighlightColorIndex = wdPink ' ���� ������ �����
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
    
   '������� ������ ���� ����� ����� ��� ������
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
    
   '������ ����������� ����� ������������ � ������� ���������
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
    
    '�������� ����������� E+0 �� �10
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Highlight = True
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Highlight = False
    With Selection.Find
        .Text = "E+"
        .Replacement.Text = "�10"
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
    
    '�������� ����������� E �� 10
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Highlight = True
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Highlight = False
    With Selection.Find
        .Text = "E"
        .Replacement.Text = "�10"
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
    
    '������ ������ �������� ������
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Highlight = False
    Selection.Find.Replacement.Text = ""
    
End Sub



Sub AutoExec()
'Sub AutoExec()
    '''''����������� ��������� ������ �� ���������
    '''''CustomizationContext = NormalTemplate
    '''''KeyBindings.ClearAll
'''If Application.CommandBars.Count > 203 Then Exit Sub

'������ ��� ������
On Error Resume Next '���������� ������
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
Application.CommandBars("����").Delete
On Error GoTo 0  '����� �� ���������� ������

  With Application.CommandBars.Add("����", tamporary = True)
     With .Controls.Add
     .Caption = "�������� � ���������� ��������� ���."
     .FaceId = 2063
     .Style = 3
     .TooltipText = "�������� � ���������� ��������� ���. �� ������� ����� ����������� ���������"
     .OnAction = "�������_�_����������"
     End With
     .Visible = True
  End With
  
'�������� ���� � ������� ���. � ��������.
  With Application.CommandBars.Add("����", tamporary = True)
     With .Controls.Add
     .Caption = "�������� ���� � ������� ���. � ��������."
     .FaceId = 9419 '3145
     .Style = 3
     .TooltipText = ""
     .OnAction = "���������_����_�_�������_�_����������"
     End With
     .Visible = True
  End With
  
  With Application.CommandBars.Add("����", tamporary = True)
     With .Controls.Add
     .Caption = "������ � ����������� ��� ���."
     .FaceId = 303
     .Style = 3
     .TooltipText = "������ � ����������� ��� ���. �� ""����� ����������� ���������"" (������, ��������� �� ��������������; ����� ��� ������ ������� ������, ������� �������������� ������� �������� ����� �� �������). �� �������� ������� ������ ������ ����� ���� (����:3.1-3.3)."
     .OnAction = "��������_��������_�_������_����_������"
     End With
     .Visible = True
  End With
  
  With Application.CommandBars.Add("����", tamporary = True)
     With .Controls.Add
     .Caption = "�������/�� ������� ����������, ������ �����"
     .FaceId = 1382
     .Style = 3
     .TooltipText = "Ctrl+W �������/�� ������� ����������, ������ ���������� �����"
     .OnAction = "����������"
     End With
     .Visible = True
  End With
  
  With Application.CommandBars.Add("����", tamporary = True)
     With .Controls.Add
     .Caption = "��������� ����� �� ������ ������ ����."
     .FaceId = 1355 '542
     .Style = 3
     .TooltipText = "Ctrl+E ����������� ����� � ���������� ������ �� ������"
     .OnAction = "��������"
     End With
     .Visible = True
  End With
  
  With Application.CommandBars.Add("����", tamporary = True)
     With .Controls.Add
     .Caption = "��������� � ���� TXT ������ ������ ������"
     .FaceId = 1548 '139
     .Style = 3
     On Error Resume Next '���������� ������
     .TooltipText = "���� TXT ����������� � ��� �� �����, ��� � ���� ��������� ��� ������ """ & ActiveDocument.Name & "_��������.txt"""
     On Error GoTo 0  '����� �� ���������� ������
     .OnAction = "��������_������_���_�_����"
     End With
     .Visible = True
  End With
  
  With Application.CommandBars.Add("����", tamporary = True)
     With .Controls.Add
     .Caption = "������ ������� c ����������������� ������������"
     .FaceId = 1446
     .Style = 3
     .TooltipText = "�������� ������ ������� � ������� ����������� ���������� �� ��������� � ""����� ����������� ���������"""
     .OnAction = "��������_���������_������������_�_������_�_���������_�_�������"
     End With
     .Visible = True
  End With
  
  With Application.CommandBars.Add("����", tamporary = True)
     With .Controls.Add
     .Caption = "������� ����������� �������� � ������������"
     .FaceId = 3202 '214'1716
     .Style = 3
     .TooltipText = "������� ����������� �������� � ������������ (����� � ������� ���������� �� �������). ������ ������� � TXT ������ ���� ����������."
     .OnAction = "��������_�������_�������"
     End With
     .Visible = True
  End With
  
  With Application.CommandBars.Add("����", tamporary = True)
     With .Controls.Add
     .Caption = "�������"
     .FaceId = 984 '214'1716
     .Style = 3
     .TooltipText = ""
     .OnAction = "����������"
     End With
     .Visible = True
  End With
  
'������_�_�������
  With Application.CommandBars.Add("����", tamporary = True)
     With .Controls.Add
     .Caption = "������ � �������"
     .FaceId = 6735 '394
     .Style = 3
     .TooltipText = "Ctrl+1 ��������� ����� ������ � ������� ���������"
     .OnAction = "������_�_�������"
     End With
     .Visible = True
  End With
  
'������_�_������_1
  With Application.CommandBars.Add("����", tamporary = True)
     With .Controls.Add
     .Caption = "������ � ������"
     .FaceId = 6751 '351
     .Style = 3
     .TooltipText = "Ctrl+2 ��������� ����� ������ � ������ ���������"
     .OnAction = "������_�_������_1"
     End With
     .Visible = True
  End With
  
'������_�_�������
  With Application.CommandBars.Add("����", tamporary = True)
     With .Controls.Add
     .Caption = "������ � �������"
     .FaceId = 6743 '352
     .Style = 3
     .TooltipText = "Ctrl+3 ��������� ����� ������ � ������� ���������"
     .OnAction = "������_�_�������"
     End With
     .Visible = True
  End With
  
'�������� ������� �� �������
  With Application.CommandBars.Add("����", tamporary = True)
     With .Controls.Add
     .Caption = "������ ����� �� �������"
     .FaceId = 382 '
     .Style = 3
     .TooltipText = "�������� ����� � �������, ������� ������� � ������� ������������ � ��������: ���, ���������� � ���. �����"
     .OnAction = "��������_������_�����_��_�������"
     End With
     .Visible = True
  End With
  
'������ ��������� ��� � � ����������� �� ����� ����� �� ��������� ���� (������ ������� �����)
  With Application.CommandBars.Add("����", tamporary = True)
     With .Controls.Add
     .Caption = "������ � ����������� ��������� ���."
     .FaceId = 159 '
     .Style = 3
     .TooltipText = "������ ��������� ��� �� ���� �������. �����. �� ��� ����� �������� � ���� ����"
     .OnAction = "��������_������_�_�����������_������_���������_���������"
     End With
     .Visible = True
  End With
  
'������� �� ����� ����������
  With Application.CommandBars.Add("����", tamporary = True)
     With .Controls.Add
     .Caption = "������� �� ����� ����������"
     .FaceId = 4305 '330 1088
     .Style = 3
     .TooltipText = "������� �� ����� ����������"
     .OnAction = "�������_�����"
     End With
     .Visible = True
  End With

'����������� ���. �����. �� ���� ������
  With Application.CommandBars.Add("����", tamporary = True)
     With .Controls.Add
     .Caption = "����������� ���. ����� �� ���� ������"
     .FaceId = 283 '
     .Style = 3
     .TooltipText = "����������� ����� ����� �� ���� ������ ���������"
     .OnAction = "��������_��������_����_�������"
     End With
     .Visible = True
  End With
  
'���������� � ������
  With Application.CommandBars.Add("����", tamporary = True)
     With .Controls.Add
     .Caption = "���������� � ������ ����."
     .FaceId = 50 '
     .Style = 3
     .TooltipText = "Ctrl+R ���������� � ������ ����."
     .OnAction = "����������_�_������"
     End With
     .Visible = True
  End With
  
'���������� ����� ����� � ������
  With Application.CommandBars.Add("����", tamporary = True)
     With .Controls.Add
     .Caption = "���������� ����� ����� � ������"
     .FaceId = 385
     .Style = 3
     .TooltipText = "Ctrl+Q ����� ��� ��� ������ �� ������ - ������� �������� � ������ ��� ���� �������� ����� ����� ������� (������������ ���������� �� ����� �������)"""
     .OnAction = "����_�����"
     End With
     .Visible = True
  End With
  
'������ ����������
  With Application.CommandBars.Add("����", tamporary = True)
     With .Controls.Add
     .Caption = "������ ����������"
     .FaceId = 1031 '
     .Style = 3
     .TooltipText = "������ ���������� �������� � ����� ������"
     .OnAction = "�_����������"
     End With
     .Visible = True
  End With
  
'��������� �����
  With Application.CommandBars.Add("����", tamporary = True)
     With .Controls.Add
     .Caption = "��������� �����"
     .FaceId = 1922 '202
     .Style = 3
     .TooltipText = "������ ����� ���������� ���"
     .OnAction = "��������_��������_�����_������"
     End With
     .Visible = True
  End With
  
'������ ����
  With Application.CommandBars.Add("����", tamporary = True)
     With .Controls.Add
     .Caption = "��������� �������"
     .FaceId = 2233 '635
     .Style = 3
     .TooltipText = "������ ������ �� ���� ��� �� ����� ���������"
     .OnAction = "������_����"
     End With
     .Visible = True
  End With
  
'����� ����������
  With Application.CommandBars.Add("����", tamporary = True)
     With .Controls.Add
     .Caption = "����� ����������"
     .FaceId = 57 '
     .Style = 3
     .TooltipText = "�������� �+3 �� 10^3"
     .OnAction = "�����_����������"
     End With
     .Visible = True
  End With
  
  Application.CommandBars.Add("����", tamporary = True).Controls.Add
    
    '��������� ������� �������
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyQ, wdKeyControl), _
        KeyCategory:=wdKeyCategoryCommand, Command:="����_�����"
        
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyE, wdKeyControl), _
        KeyCategory:=wdKeyCategoryCommand, Command:="��������"
        
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyW, wdKeyControl), _
        KeyCategory:=wdKeyCategoryCommand, Command:="����������"
    '������_�_�������
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKey1, wdKeyControl), _
        KeyCategory:=wdKeyCategoryCommand, Command:="������_�_�������"
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyNumeric1, wdKeyControl), _
        KeyCategory:=wdKeyCategoryCommand, Command:="������_�_�������"
    '������_�_������_1
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKey2, wdKeyControl), _
        KeyCategory:=wdKeyCategoryCommand, Command:="������_�_������_1"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyNumeric2, wdKeyControl), _
        KeyCategory:=wdKeyCategoryCommand, Command:="������_�_������_1"
    '������_�_�������
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKey3, wdKeyControl), _
        KeyCategory:=wdKeyCategoryCommand, Command:="������_�_�������"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyNumeric3, wdKeyControl), _
        KeyCategory:=wdKeyCategoryCommand, Command:="������_�_�������"
    '������_�_�������
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKey4, wdKeyControl), _
        KeyCategory:=wdKeyCategoryCommand, Command:="������_�_�������"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyNumeric4, wdKeyControl), _
        KeyCategory:=wdKeyCategoryCommand, Command:="������_�_�������"
    '������_�_�����
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKey5, wdKeyControl), _
        KeyCategory:=wdKeyCategoryCommand, Command:="������_�_�����"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyNumeric5, wdKeyControl), _
        KeyCategory:=wdKeyCategoryCommand, Command:="������_�_�����"
    '���������� � ������
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyR, wdKeyControl), _
        KeyCategory:=wdKeyCategoryCommand, Command:="����������_�_������"
    
End Sub


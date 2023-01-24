VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Загрузка"
   ClientHeight    =   1080
   ClientLeft      =   48
   ClientTop       =   376
   ClientWidth     =   7080
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton1_Click()
If ПОЛНЫЙ_поиск = True Then
   Set objshell = Nothing
   Set ie = Nothing
   Set СписокГОСТов = Nothing
   ie.Quit
End If
Stop
Application.ScreenUpdating = False 'выключить обновление документа
End Sub

Sub UserForm_Activate()
'On Error GoTo cancelhandler
'Application.EnableCancelKey = xlerrorhandler
'Application.EnableCancelKey = xlDisabled
If ГАЛКА_Проверка_и_замена_всех_колонт = True Then Проверка_и_замена_всех_колонт
If ГАЛКА_Список_стр_в_файл = True Then Список_стр_в_файл
If ГАЛКА_Сравнение_колонтитулов_в_листах_с_указаными_в_таблице = True Then Сравнение_колонтитулов_в_листах_с_указаными_в_таблице
If ГАЛКА_Удаляем_разрывы = True Then Удаляем_разрывы
If ГАЛКА_Замена_точек_на_запятые = True Then Замена_точек_на_запятые
If ГАЛКА_Вносим_в_колонтитулы_только_последние_изменения = True Then Вносим_в_колонтитулы_только_последние_изменения
If ГАЛКА_Пересчет_всех_пунктов = True Then Пересчет_всех_пунктов
If ГАЛКА_Проверка_годов_ГОСТов = True Then Проверка_годов_ГОСТов

'cancelhandler:
'Application.EnableCancelKey = xlInterrupt
'If Err.Number = 18 Then MsgBox "5"
'Application.ScreenUpdating = True 'включить обновление документа1.Hide
'Application.EnableCancelKey = xlInterrupt
'UserForm1.Repaint

Unload UserForm1 'забыть форму и все данные на ней
'UserForm1.Hide  'скрыть форму данные на ней не забубутся
End Sub


'Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'Stop
'End Sub

Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii = 27 Then Stop 'esc
Application.ScreenUpdating = False 'выключить обновление документа
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = 0 Then Stop 'если закрыли руками
Application.ScreenUpdating = False 'выключить обновление документа
End Sub

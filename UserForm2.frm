VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Список листов"
   ClientHeight    =   10260
   ClientLeft      =   48
   ClientTop       =   376
   ClientWidth     =   11112
   OleObjectBlob   =   "UserForm2.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public не_меняем_размер_формы As Boolean

Private Sub CommandButton1_Click()
On Error Resume Next 'Пропустить все ошибки
     Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=CInt(UserForm2.TextBox3.Text), Name:="" 'перейти на стр по номеру
End Sub

Private Sub CommandButton2_Click()
 On Error Resume Next 'Пропустить все ошибки
     Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=ActiveDocument.ComputeStatistics(wdStatisticPages), Name:=""  'перейти на стр по номеру
End Sub



Private Sub Label1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
On Error Resume Next 'пропускаем ошибки
If Button = 2 Then
  Ширена_формы = UserForm2.Width
  'Растояние_между_зам_и_нов = TextBox_а1_нов.Left - (TextBox_а1_зам.Left + TextBox_а1_зам.Width)
  'Растояние_между_нов_и_не = TextBox_а1_не.Left - (TextBox_а1_нов.Left + TextBox_а1_нов.Width)
  

  UserForm2.Height = Y + 22
  UserForm2.Width = x + 4
  
  TextBox_а1_зам.Width = UserForm2.Width * TextBox_а1_зам.Width / Ширена_формы
  TextBox_а1_нов.Width = UserForm2.Width * TextBox_а1_нов.Width / Ширена_формы
  TextBox_а1_нов.Left = UserForm2.Width * TextBox_а1_нов.Left / Ширена_формы
  TextBox_а1_не.Width = UserForm2.Width * TextBox_а1_не.Width / Ширена_формы
  TextBox_а1_не.Left = UserForm2.Width * TextBox_а1_не.Left / Ширена_формы
  
  TextBox_а2_зам.Width = UserForm2.Width * TextBox_а2_зам.Width / Ширена_формы
  TextBox_а2_нов.Width = UserForm2.Width * TextBox_а2_нов.Width / Ширена_формы
  TextBox_а2_нов.Left = UserForm2.Width * TextBox_а2_нов.Left / Ширена_формы
  TextBox_а2_не.Width = UserForm2.Width * TextBox_а2_не.Width / Ширена_формы
  TextBox_а2_не.Left = UserForm2.Width * TextBox_а2_не.Left / Ширена_формы
  
  TextBox_а3_зам.Width = UserForm2.Width * TextBox_а3_зам.Width / Ширена_формы
  TextBox_а3_нов.Width = UserForm2.Width * TextBox_а3_нов.Width / Ширена_формы
  TextBox_а3_нов.Left = UserForm2.Width * TextBox_а3_нов.Left / Ширена_формы
  TextBox_а3_не.Width = UserForm2.Width * TextBox_а3_не.Width / Ширена_формы
  TextBox_а3_не.Left = UserForm2.Width * TextBox_а3_не.Left / Ширена_формы
  
  TextBox_а4_зам.Width = UserForm2.Width * TextBox_а4_зам.Width / Ширена_формы
  TextBox_а4_нов.Width = UserForm2.Width * TextBox_а4_нов.Width / Ширена_формы
  TextBox_а4_нов.Left = UserForm2.Width * TextBox_а4_нов.Left / Ширена_формы
  TextBox_а4_не.Width = UserForm2.Width * TextBox_а4_не.Width / Ширена_формы
  TextBox_а4_не.Left = UserForm2.Width * TextBox_а4_не.Left / Ширена_формы
  
  TextBox_а5_зам.Width = UserForm2.Width * TextBox_а5_зам.Width / Ширена_формы
  TextBox_а5_нов.Width = UserForm2.Width * TextBox_а5_нов.Width / Ширена_формы
  TextBox_а5_нов.Left = UserForm2.Width * TextBox_а5_нов.Left / Ширена_формы
  TextBox_а5_не.Width = UserForm2.Width * TextBox_а5_не.Width / Ширена_формы
  TextBox_а5_не.Left = UserForm2.Width * TextBox_а5_не.Left / Ширена_формы
  
  TextBox_а6_зам.Width = UserForm2.Width * TextBox_а6_зам.Width / Ширена_формы
  TextBox_а6_нов.Width = UserForm2.Width * TextBox_а6_нов.Width / Ширена_формы
  TextBox_а6_нов.Left = UserForm2.Width * TextBox_а6_нов.Left / Ширена_формы
  TextBox_а6_не.Width = UserForm2.Width * TextBox_а6_не.Width / Ширена_формы
  TextBox_а6_не.Left = UserForm2.Width * TextBox_а6_не.Left / Ширена_формы
  
  TextBox_а7_зам.Width = UserForm2.Width * TextBox_а7_зам.Width / Ширена_формы
  TextBox_а7_нов.Width = UserForm2.Width * TextBox_а7_нов.Width / Ширена_формы
  TextBox_а7_нов.Left = UserForm2.Width * TextBox_а7_нов.Left / Ширена_формы
  TextBox_а7_не.Width = UserForm2.Width * TextBox_а7_не.Width / Ширена_формы
  TextBox_а7_не.Left = UserForm2.Width * TextBox_а7_не.Left / Ширена_формы
  
  TextBox_а8_зам.Width = UserForm2.Width * TextBox_а8_зам.Width / Ширена_формы
  TextBox_а8_нов.Width = UserForm2.Width * TextBox_а8_нов.Width / Ширена_формы
  TextBox_а8_нов.Left = UserForm2.Width * TextBox_а8_нов.Left / Ширена_формы
  TextBox_а8_не.Width = UserForm2.Width * TextBox_а8_не.Width / Ширена_формы
  TextBox_а8_не.Left = UserForm2.Width * TextBox_а8_не.Left / Ширена_формы
  
  TextBox_а9_зам.Width = UserForm2.Width * TextBox_а9_зам.Width / Ширена_формы
  TextBox_а9_нов.Width = UserForm2.Width * TextBox_а9_нов.Width / Ширена_формы
  TextBox_а9_нов.Left = UserForm2.Width * TextBox_а9_нов.Left / Ширена_формы
  TextBox_а9_не.Width = UserForm2.Width * TextBox_а9_не.Width / Ширена_формы
  TextBox_а9_не.Left = UserForm2.Width * TextBox_а9_не.Left / Ширена_формы
  
  TextBox_а10_зам.Width = UserForm2.Width * TextBox_а10_зам.Width / Ширена_формы
  TextBox_а10_нов.Width = UserForm2.Width * TextBox_а10_нов.Width / Ширена_формы
  TextBox_а10_нов.Left = UserForm2.Width * TextBox_а10_нов.Left / Ширена_формы
  TextBox_а10_не.Width = UserForm2.Width * TextBox_а10_не.Width / Ширена_формы
  TextBox_а10_не.Left = UserForm2.Width * TextBox_а10_не.Left / Ширена_формы
  
  TextBox_а11_зам.Width = UserForm2.Width * TextBox_а11_зам.Width / Ширена_формы
  TextBox_а11_нов.Width = UserForm2.Width * TextBox_а11_нов.Width / Ширена_формы
  TextBox_а11_нов.Left = UserForm2.Width * TextBox_а11_нов.Left / Ширена_формы
  TextBox_а11_не.Width = UserForm2.Width * TextBox_а11_не.Width / Ширена_формы
  TextBox_а11_не.Left = UserForm2.Width * TextBox_а11_не.Left / Ширена_формы
  
  TextBox_а12_зам.Width = UserForm2.Width * TextBox_а12_зам.Width / Ширена_формы
  TextBox_а12_нов.Width = UserForm2.Width * TextBox_а12_нов.Width / Ширена_формы
  TextBox_а12_нов.Left = UserForm2.Width * TextBox_а12_нов.Left / Ширена_формы
  TextBox_а12_не.Width = UserForm2.Width * TextBox_а12_не.Width / Ширена_формы
  TextBox_а12_не.Left = UserForm2.Width * TextBox_а12_не.Left / Ширена_формы
  
  TextBox_не_зам.Width = UserForm2.Width * TextBox_не_зам.Width / Ширена_формы
  TextBox4_не_нов.Width = UserForm2.Width * TextBox4_не_нов.Width / Ширена_формы
  TextBox4_не_нов.Left = UserForm2.Width * TextBox4_не_нов.Left / Ширена_формы
  TextBox_не.Width = UserForm2.Width * TextBox_не.Width / Ширена_формы
  TextBox_не.Left = UserForm2.Width * TextBox_не.Left / Ширена_формы
  
  TextBox5.Width = UserForm2.Width * TextBox5.Width / Ширена_формы
  TextBox_без_изм.Width = UserForm2.Width * TextBox_без_изм.Width / Ширена_формы
  
  Label5.Left = TextBox_а1_зам.Left + TextBox_а1_зам.Width / 2 - Label5.Width / 2
  Label6.Left = TextBox_а1_нов.Left + TextBox_а1_нов.Width / 2 - Label6.Width / 2
  Label23.Left = TextBox_а1_не.Left + TextBox_а1_не.Width / 2 - Label23.Width / 2
End If
End Sub

Private Sub UserForm_Activate()
'не_меняем_размер_формы =
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
On Error Resume Next 'пропускаем ошибки
If Button = 2 Then
  Ширена_формы = UserForm2.Width
  'Растояние_между_зам_и_нов = TextBox_а1_нов.Left - (TextBox_а1_зам.Left + TextBox_а1_зам.Width)
  'Растояние_между_нов_и_не = TextBox_а1_не.Left - (TextBox_а1_нов.Left + TextBox_а1_нов.Width)
  

  UserForm2.Height = Y + 22
  UserForm2.Width = x + 4
  
  TextBox_а1_зам.Width = UserForm2.Width * TextBox_а1_зам.Width / Ширена_формы
  TextBox_а1_нов.Width = UserForm2.Width * TextBox_а1_нов.Width / Ширена_формы
  TextBox_а1_нов.Left = UserForm2.Width * TextBox_а1_нов.Left / Ширена_формы
  TextBox_а1_не.Width = UserForm2.Width * TextBox_а1_не.Width / Ширена_формы
  TextBox_а1_не.Left = UserForm2.Width * TextBox_а1_не.Left / Ширена_формы
  
  TextBox_а2_зам.Width = UserForm2.Width * TextBox_а2_зам.Width / Ширена_формы
  TextBox_а2_нов.Width = UserForm2.Width * TextBox_а2_нов.Width / Ширена_формы
  TextBox_а2_нов.Left = UserForm2.Width * TextBox_а2_нов.Left / Ширена_формы
  TextBox_а2_не.Width = UserForm2.Width * TextBox_а2_не.Width / Ширена_формы
  TextBox_а2_не.Left = UserForm2.Width * TextBox_а2_не.Left / Ширена_формы
  
  TextBox_а3_зам.Width = UserForm2.Width * TextBox_а3_зам.Width / Ширена_формы
  TextBox_а3_нов.Width = UserForm2.Width * TextBox_а3_нов.Width / Ширена_формы
  TextBox_а3_нов.Left = UserForm2.Width * TextBox_а3_нов.Left / Ширена_формы
  TextBox_а3_не.Width = UserForm2.Width * TextBox_а3_не.Width / Ширена_формы
  TextBox_а3_не.Left = UserForm2.Width * TextBox_а3_не.Left / Ширена_формы
  
  TextBox_а4_зам.Width = UserForm2.Width * TextBox_а4_зам.Width / Ширена_формы
  TextBox_а4_нов.Width = UserForm2.Width * TextBox_а4_нов.Width / Ширена_формы
  TextBox_а4_нов.Left = UserForm2.Width * TextBox_а4_нов.Left / Ширена_формы
  TextBox_а4_не.Width = UserForm2.Width * TextBox_а4_не.Width / Ширена_формы
  TextBox_а4_не.Left = UserForm2.Width * TextBox_а4_не.Left / Ширена_формы
  
  TextBox_а5_зам.Width = UserForm2.Width * TextBox_а5_зам.Width / Ширена_формы
  TextBox_а5_нов.Width = UserForm2.Width * TextBox_а5_нов.Width / Ширена_формы
  TextBox_а5_нов.Left = UserForm2.Width * TextBox_а5_нов.Left / Ширена_формы
  TextBox_а5_не.Width = UserForm2.Width * TextBox_а5_не.Width / Ширена_формы
  TextBox_а5_не.Left = UserForm2.Width * TextBox_а5_не.Left / Ширена_формы
  
  TextBox_а6_зам.Width = UserForm2.Width * TextBox_а6_зам.Width / Ширена_формы
  TextBox_а6_нов.Width = UserForm2.Width * TextBox_а6_нов.Width / Ширена_формы
  TextBox_а6_нов.Left = UserForm2.Width * TextBox_а6_нов.Left / Ширена_формы
  TextBox_а6_не.Width = UserForm2.Width * TextBox_а6_не.Width / Ширена_формы
  TextBox_а6_не.Left = UserForm2.Width * TextBox_а6_не.Left / Ширена_формы
  
  TextBox_а7_зам.Width = UserForm2.Width * TextBox_а7_зам.Width / Ширена_формы
  TextBox_а7_нов.Width = UserForm2.Width * TextBox_а7_нов.Width / Ширена_формы
  TextBox_а7_нов.Left = UserForm2.Width * TextBox_а7_нов.Left / Ширена_формы
  TextBox_а7_не.Width = UserForm2.Width * TextBox_а7_не.Width / Ширена_формы
  TextBox_а7_не.Left = UserForm2.Width * TextBox_а7_не.Left / Ширена_формы
  
  TextBox_а8_зам.Width = UserForm2.Width * TextBox_а8_зам.Width / Ширена_формы
  TextBox_а8_нов.Width = UserForm2.Width * TextBox_а8_нов.Width / Ширена_формы
  TextBox_а8_нов.Left = UserForm2.Width * TextBox_а8_нов.Left / Ширена_формы
  TextBox_а8_не.Width = UserForm2.Width * TextBox_а8_не.Width / Ширена_формы
  TextBox_а8_не.Left = UserForm2.Width * TextBox_а8_не.Left / Ширена_формы
  
  TextBox_а9_зам.Width = UserForm2.Width * TextBox_а9_зам.Width / Ширена_формы
  TextBox_а9_нов.Width = UserForm2.Width * TextBox_а9_нов.Width / Ширена_формы
  TextBox_а9_нов.Left = UserForm2.Width * TextBox_а9_нов.Left / Ширена_формы
  TextBox_а9_не.Width = UserForm2.Width * TextBox_а9_не.Width / Ширена_формы
  TextBox_а9_не.Left = UserForm2.Width * TextBox_а9_не.Left / Ширена_формы
  
  TextBox_а10_зам.Width = UserForm2.Width * TextBox_а10_зам.Width / Ширена_формы
  TextBox_а10_нов.Width = UserForm2.Width * TextBox_а10_нов.Width / Ширена_формы
  TextBox_а10_нов.Left = UserForm2.Width * TextBox_а10_нов.Left / Ширена_формы
  TextBox_а10_не.Width = UserForm2.Width * TextBox_а10_не.Width / Ширена_формы
  TextBox_а10_не.Left = UserForm2.Width * TextBox_а10_не.Left / Ширена_формы
  
  TextBox_а11_зам.Width = UserForm2.Width * TextBox_а11_зам.Width / Ширена_формы
  TextBox_а11_нов.Width = UserForm2.Width * TextBox_а11_нов.Width / Ширена_формы
  TextBox_а11_нов.Left = UserForm2.Width * TextBox_а11_нов.Left / Ширена_формы
  TextBox_а11_не.Width = UserForm2.Width * TextBox_а11_не.Width / Ширена_формы
  TextBox_а11_не.Left = UserForm2.Width * TextBox_а11_не.Left / Ширена_формы
  
  TextBox_а12_зам.Width = UserForm2.Width * TextBox_а12_зам.Width / Ширена_формы
  TextBox_а12_нов.Width = UserForm2.Width * TextBox_а12_нов.Width / Ширена_формы
  TextBox_а12_нов.Left = UserForm2.Width * TextBox_а12_нов.Left / Ширена_формы
  TextBox_а12_не.Width = UserForm2.Width * TextBox_а12_не.Width / Ширена_формы
  TextBox_а12_не.Left = UserForm2.Width * TextBox_а12_не.Left / Ширена_формы
  
  TextBox_не_зам.Width = UserForm2.Width * TextBox_не_зам.Width / Ширена_формы
  TextBox4_не_нов.Width = UserForm2.Width * TextBox4_не_нов.Width / Ширена_формы
  TextBox4_не_нов.Left = UserForm2.Width * TextBox4_не_нов.Left / Ширена_формы
  TextBox_не.Width = UserForm2.Width * TextBox_не.Width / Ширена_формы
  TextBox_не.Left = UserForm2.Width * TextBox_не.Left / Ширена_формы
  
  TextBox5.Width = UserForm2.Width * TextBox5.Width / Ширена_формы
  TextBox_без_изм.Width = UserForm2.Width * TextBox_без_изм.Width / Ширена_формы
  
  Label5.Left = TextBox_а1_зам.Left + TextBox_а1_зам.Width / 2 - Label5.Width / 2
  Label6.Left = TextBox_а1_нов.Left + TextBox_а1_нов.Width / 2 - Label6.Width / 2
  Label23.Left = TextBox_а1_не.Left + TextBox_а1_не.Width / 2 - Label23.Width / 2
  
  'TextBox_а1_зам.Width = TextBox_а1_зам.Width + CInt((UserForm2.Width - Ширена_формы) / 3)
  'TextBox_а1_нов.Width = TextBox_а1_нов.Width + CInt((UserForm2.Width - Ширена_формы) / 3)
  'TextBox_а1_нов.Left = TextBox_а1_зам.Left + TextBox_а1_зам.Width + Растояние_между_зам_и_нов
  'TextBox_а1_не.Width = TextBox_а1_не.Width + CInt((UserForm2.Width - Ширена_формы) / 3)
  'TextBox_а1_не.Left = TextBox_а1_нов.Left + TextBox_а1_нов.Width + Растояние_между_нов_и_не
  


  
'  не_меняем_размер_формы = True
End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'UserForm2.Hide
End Sub

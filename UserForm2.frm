VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "������ ������"
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


Public ��_������_������_����� As Boolean

Private Sub CommandButton1_Click()
On Error Resume Next '���������� ��� ������
     Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=CInt(UserForm2.TextBox3.Text), Name:="" '������� �� ��� �� ������
End Sub

Private Sub CommandButton2_Click()
 On Error Resume Next '���������� ��� ������
     Selection.GoTo What:=wdGoToPage, Which:=wdGoToFirst, Count:=ActiveDocument.ComputeStatistics(wdStatisticPages), Name:=""  '������� �� ��� �� ������
End Sub



Private Sub Label1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
On Error Resume Next '���������� ������
If Button = 2 Then
  ������_����� = UserForm2.Width
  '���������_�����_���_�_��� = TextBox_�1_���.Left - (TextBox_�1_���.Left + TextBox_�1_���.Width)
  '���������_�����_���_�_�� = TextBox_�1_��.Left - (TextBox_�1_���.Left + TextBox_�1_���.Width)
  

  UserForm2.Height = Y + 22
  UserForm2.Width = x + 4
  
  TextBox_�1_���.Width = UserForm2.Width * TextBox_�1_���.Width / ������_�����
  TextBox_�1_���.Width = UserForm2.Width * TextBox_�1_���.Width / ������_�����
  TextBox_�1_���.Left = UserForm2.Width * TextBox_�1_���.Left / ������_�����
  TextBox_�1_��.Width = UserForm2.Width * TextBox_�1_��.Width / ������_�����
  TextBox_�1_��.Left = UserForm2.Width * TextBox_�1_��.Left / ������_�����
  
  TextBox_�2_���.Width = UserForm2.Width * TextBox_�2_���.Width / ������_�����
  TextBox_�2_���.Width = UserForm2.Width * TextBox_�2_���.Width / ������_�����
  TextBox_�2_���.Left = UserForm2.Width * TextBox_�2_���.Left / ������_�����
  TextBox_�2_��.Width = UserForm2.Width * TextBox_�2_��.Width / ������_�����
  TextBox_�2_��.Left = UserForm2.Width * TextBox_�2_��.Left / ������_�����
  
  TextBox_�3_���.Width = UserForm2.Width * TextBox_�3_���.Width / ������_�����
  TextBox_�3_���.Width = UserForm2.Width * TextBox_�3_���.Width / ������_�����
  TextBox_�3_���.Left = UserForm2.Width * TextBox_�3_���.Left / ������_�����
  TextBox_�3_��.Width = UserForm2.Width * TextBox_�3_��.Width / ������_�����
  TextBox_�3_��.Left = UserForm2.Width * TextBox_�3_��.Left / ������_�����
  
  TextBox_�4_���.Width = UserForm2.Width * TextBox_�4_���.Width / ������_�����
  TextBox_�4_���.Width = UserForm2.Width * TextBox_�4_���.Width / ������_�����
  TextBox_�4_���.Left = UserForm2.Width * TextBox_�4_���.Left / ������_�����
  TextBox_�4_��.Width = UserForm2.Width * TextBox_�4_��.Width / ������_�����
  TextBox_�4_��.Left = UserForm2.Width * TextBox_�4_��.Left / ������_�����
  
  TextBox_�5_���.Width = UserForm2.Width * TextBox_�5_���.Width / ������_�����
  TextBox_�5_���.Width = UserForm2.Width * TextBox_�5_���.Width / ������_�����
  TextBox_�5_���.Left = UserForm2.Width * TextBox_�5_���.Left / ������_�����
  TextBox_�5_��.Width = UserForm2.Width * TextBox_�5_��.Width / ������_�����
  TextBox_�5_��.Left = UserForm2.Width * TextBox_�5_��.Left / ������_�����
  
  TextBox_�6_���.Width = UserForm2.Width * TextBox_�6_���.Width / ������_�����
  TextBox_�6_���.Width = UserForm2.Width * TextBox_�6_���.Width / ������_�����
  TextBox_�6_���.Left = UserForm2.Width * TextBox_�6_���.Left / ������_�����
  TextBox_�6_��.Width = UserForm2.Width * TextBox_�6_��.Width / ������_�����
  TextBox_�6_��.Left = UserForm2.Width * TextBox_�6_��.Left / ������_�����
  
  TextBox_�7_���.Width = UserForm2.Width * TextBox_�7_���.Width / ������_�����
  TextBox_�7_���.Width = UserForm2.Width * TextBox_�7_���.Width / ������_�����
  TextBox_�7_���.Left = UserForm2.Width * TextBox_�7_���.Left / ������_�����
  TextBox_�7_��.Width = UserForm2.Width * TextBox_�7_��.Width / ������_�����
  TextBox_�7_��.Left = UserForm2.Width * TextBox_�7_��.Left / ������_�����
  
  TextBox_�8_���.Width = UserForm2.Width * TextBox_�8_���.Width / ������_�����
  TextBox_�8_���.Width = UserForm2.Width * TextBox_�8_���.Width / ������_�����
  TextBox_�8_���.Left = UserForm2.Width * TextBox_�8_���.Left / ������_�����
  TextBox_�8_��.Width = UserForm2.Width * TextBox_�8_��.Width / ������_�����
  TextBox_�8_��.Left = UserForm2.Width * TextBox_�8_��.Left / ������_�����
  
  TextBox_�9_���.Width = UserForm2.Width * TextBox_�9_���.Width / ������_�����
  TextBox_�9_���.Width = UserForm2.Width * TextBox_�9_���.Width / ������_�����
  TextBox_�9_���.Left = UserForm2.Width * TextBox_�9_���.Left / ������_�����
  TextBox_�9_��.Width = UserForm2.Width * TextBox_�9_��.Width / ������_�����
  TextBox_�9_��.Left = UserForm2.Width * TextBox_�9_��.Left / ������_�����
  
  TextBox_�10_���.Width = UserForm2.Width * TextBox_�10_���.Width / ������_�����
  TextBox_�10_���.Width = UserForm2.Width * TextBox_�10_���.Width / ������_�����
  TextBox_�10_���.Left = UserForm2.Width * TextBox_�10_���.Left / ������_�����
  TextBox_�10_��.Width = UserForm2.Width * TextBox_�10_��.Width / ������_�����
  TextBox_�10_��.Left = UserForm2.Width * TextBox_�10_��.Left / ������_�����
  
  TextBox_�11_���.Width = UserForm2.Width * TextBox_�11_���.Width / ������_�����
  TextBox_�11_���.Width = UserForm2.Width * TextBox_�11_���.Width / ������_�����
  TextBox_�11_���.Left = UserForm2.Width * TextBox_�11_���.Left / ������_�����
  TextBox_�11_��.Width = UserForm2.Width * TextBox_�11_��.Width / ������_�����
  TextBox_�11_��.Left = UserForm2.Width * TextBox_�11_��.Left / ������_�����
  
  TextBox_�12_���.Width = UserForm2.Width * TextBox_�12_���.Width / ������_�����
  TextBox_�12_���.Width = UserForm2.Width * TextBox_�12_���.Width / ������_�����
  TextBox_�12_���.Left = UserForm2.Width * TextBox_�12_���.Left / ������_�����
  TextBox_�12_��.Width = UserForm2.Width * TextBox_�12_��.Width / ������_�����
  TextBox_�12_��.Left = UserForm2.Width * TextBox_�12_��.Left / ������_�����
  
  TextBox_��_���.Width = UserForm2.Width * TextBox_��_���.Width / ������_�����
  TextBox4_��_���.Width = UserForm2.Width * TextBox4_��_���.Width / ������_�����
  TextBox4_��_���.Left = UserForm2.Width * TextBox4_��_���.Left / ������_�����
  TextBox_��.Width = UserForm2.Width * TextBox_��.Width / ������_�����
  TextBox_��.Left = UserForm2.Width * TextBox_��.Left / ������_�����
  
  TextBox5.Width = UserForm2.Width * TextBox5.Width / ������_�����
  TextBox_���_���.Width = UserForm2.Width * TextBox_���_���.Width / ������_�����
  
  Label5.Left = TextBox_�1_���.Left + TextBox_�1_���.Width / 2 - Label5.Width / 2
  Label6.Left = TextBox_�1_���.Left + TextBox_�1_���.Width / 2 - Label6.Width / 2
  Label23.Left = TextBox_�1_��.Left + TextBox_�1_��.Width / 2 - Label23.Width / 2
End If
End Sub

Private Sub UserForm_Activate()
'��_������_������_����� =
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
On Error Resume Next '���������� ������
If Button = 2 Then
  ������_����� = UserForm2.Width
  '���������_�����_���_�_��� = TextBox_�1_���.Left - (TextBox_�1_���.Left + TextBox_�1_���.Width)
  '���������_�����_���_�_�� = TextBox_�1_��.Left - (TextBox_�1_���.Left + TextBox_�1_���.Width)
  

  UserForm2.Height = Y + 22
  UserForm2.Width = x + 4
  
  TextBox_�1_���.Width = UserForm2.Width * TextBox_�1_���.Width / ������_�����
  TextBox_�1_���.Width = UserForm2.Width * TextBox_�1_���.Width / ������_�����
  TextBox_�1_���.Left = UserForm2.Width * TextBox_�1_���.Left / ������_�����
  TextBox_�1_��.Width = UserForm2.Width * TextBox_�1_��.Width / ������_�����
  TextBox_�1_��.Left = UserForm2.Width * TextBox_�1_��.Left / ������_�����
  
  TextBox_�2_���.Width = UserForm2.Width * TextBox_�2_���.Width / ������_�����
  TextBox_�2_���.Width = UserForm2.Width * TextBox_�2_���.Width / ������_�����
  TextBox_�2_���.Left = UserForm2.Width * TextBox_�2_���.Left / ������_�����
  TextBox_�2_��.Width = UserForm2.Width * TextBox_�2_��.Width / ������_�����
  TextBox_�2_��.Left = UserForm2.Width * TextBox_�2_��.Left / ������_�����
  
  TextBox_�3_���.Width = UserForm2.Width * TextBox_�3_���.Width / ������_�����
  TextBox_�3_���.Width = UserForm2.Width * TextBox_�3_���.Width / ������_�����
  TextBox_�3_���.Left = UserForm2.Width * TextBox_�3_���.Left / ������_�����
  TextBox_�3_��.Width = UserForm2.Width * TextBox_�3_��.Width / ������_�����
  TextBox_�3_��.Left = UserForm2.Width * TextBox_�3_��.Left / ������_�����
  
  TextBox_�4_���.Width = UserForm2.Width * TextBox_�4_���.Width / ������_�����
  TextBox_�4_���.Width = UserForm2.Width * TextBox_�4_���.Width / ������_�����
  TextBox_�4_���.Left = UserForm2.Width * TextBox_�4_���.Left / ������_�����
  TextBox_�4_��.Width = UserForm2.Width * TextBox_�4_��.Width / ������_�����
  TextBox_�4_��.Left = UserForm2.Width * TextBox_�4_��.Left / ������_�����
  
  TextBox_�5_���.Width = UserForm2.Width * TextBox_�5_���.Width / ������_�����
  TextBox_�5_���.Width = UserForm2.Width * TextBox_�5_���.Width / ������_�����
  TextBox_�5_���.Left = UserForm2.Width * TextBox_�5_���.Left / ������_�����
  TextBox_�5_��.Width = UserForm2.Width * TextBox_�5_��.Width / ������_�����
  TextBox_�5_��.Left = UserForm2.Width * TextBox_�5_��.Left / ������_�����
  
  TextBox_�6_���.Width = UserForm2.Width * TextBox_�6_���.Width / ������_�����
  TextBox_�6_���.Width = UserForm2.Width * TextBox_�6_���.Width / ������_�����
  TextBox_�6_���.Left = UserForm2.Width * TextBox_�6_���.Left / ������_�����
  TextBox_�6_��.Width = UserForm2.Width * TextBox_�6_��.Width / ������_�����
  TextBox_�6_��.Left = UserForm2.Width * TextBox_�6_��.Left / ������_�����
  
  TextBox_�7_���.Width = UserForm2.Width * TextBox_�7_���.Width / ������_�����
  TextBox_�7_���.Width = UserForm2.Width * TextBox_�7_���.Width / ������_�����
  TextBox_�7_���.Left = UserForm2.Width * TextBox_�7_���.Left / ������_�����
  TextBox_�7_��.Width = UserForm2.Width * TextBox_�7_��.Width / ������_�����
  TextBox_�7_��.Left = UserForm2.Width * TextBox_�7_��.Left / ������_�����
  
  TextBox_�8_���.Width = UserForm2.Width * TextBox_�8_���.Width / ������_�����
  TextBox_�8_���.Width = UserForm2.Width * TextBox_�8_���.Width / ������_�����
  TextBox_�8_���.Left = UserForm2.Width * TextBox_�8_���.Left / ������_�����
  TextBox_�8_��.Width = UserForm2.Width * TextBox_�8_��.Width / ������_�����
  TextBox_�8_��.Left = UserForm2.Width * TextBox_�8_��.Left / ������_�����
  
  TextBox_�9_���.Width = UserForm2.Width * TextBox_�9_���.Width / ������_�����
  TextBox_�9_���.Width = UserForm2.Width * TextBox_�9_���.Width / ������_�����
  TextBox_�9_���.Left = UserForm2.Width * TextBox_�9_���.Left / ������_�����
  TextBox_�9_��.Width = UserForm2.Width * TextBox_�9_��.Width / ������_�����
  TextBox_�9_��.Left = UserForm2.Width * TextBox_�9_��.Left / ������_�����
  
  TextBox_�10_���.Width = UserForm2.Width * TextBox_�10_���.Width / ������_�����
  TextBox_�10_���.Width = UserForm2.Width * TextBox_�10_���.Width / ������_�����
  TextBox_�10_���.Left = UserForm2.Width * TextBox_�10_���.Left / ������_�����
  TextBox_�10_��.Width = UserForm2.Width * TextBox_�10_��.Width / ������_�����
  TextBox_�10_��.Left = UserForm2.Width * TextBox_�10_��.Left / ������_�����
  
  TextBox_�11_���.Width = UserForm2.Width * TextBox_�11_���.Width / ������_�����
  TextBox_�11_���.Width = UserForm2.Width * TextBox_�11_���.Width / ������_�����
  TextBox_�11_���.Left = UserForm2.Width * TextBox_�11_���.Left / ������_�����
  TextBox_�11_��.Width = UserForm2.Width * TextBox_�11_��.Width / ������_�����
  TextBox_�11_��.Left = UserForm2.Width * TextBox_�11_��.Left / ������_�����
  
  TextBox_�12_���.Width = UserForm2.Width * TextBox_�12_���.Width / ������_�����
  TextBox_�12_���.Width = UserForm2.Width * TextBox_�12_���.Width / ������_�����
  TextBox_�12_���.Left = UserForm2.Width * TextBox_�12_���.Left / ������_�����
  TextBox_�12_��.Width = UserForm2.Width * TextBox_�12_��.Width / ������_�����
  TextBox_�12_��.Left = UserForm2.Width * TextBox_�12_��.Left / ������_�����
  
  TextBox_��_���.Width = UserForm2.Width * TextBox_��_���.Width / ������_�����
  TextBox4_��_���.Width = UserForm2.Width * TextBox4_��_���.Width / ������_�����
  TextBox4_��_���.Left = UserForm2.Width * TextBox4_��_���.Left / ������_�����
  TextBox_��.Width = UserForm2.Width * TextBox_��.Width / ������_�����
  TextBox_��.Left = UserForm2.Width * TextBox_��.Left / ������_�����
  
  TextBox5.Width = UserForm2.Width * TextBox5.Width / ������_�����
  TextBox_���_���.Width = UserForm2.Width * TextBox_���_���.Width / ������_�����
  
  Label5.Left = TextBox_�1_���.Left + TextBox_�1_���.Width / 2 - Label5.Width / 2
  Label6.Left = TextBox_�1_���.Left + TextBox_�1_���.Width / 2 - Label6.Width / 2
  Label23.Left = TextBox_�1_��.Left + TextBox_�1_��.Width / 2 - Label23.Width / 2
  
  'TextBox_�1_���.Width = TextBox_�1_���.Width + CInt((UserForm2.Width - ������_�����) / 3)
  'TextBox_�1_���.Width = TextBox_�1_���.Width + CInt((UserForm2.Width - ������_�����) / 3)
  'TextBox_�1_���.Left = TextBox_�1_���.Left + TextBox_�1_���.Width + ���������_�����_���_�_���
  'TextBox_�1_��.Width = TextBox_�1_��.Width + CInt((UserForm2.Width - ������_�����) / 3)
  'TextBox_�1_��.Left = TextBox_�1_���.Left + TextBox_�1_���.Width + ���������_�����_���_�_��
  


  
'  ��_������_������_����� = True
End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'UserForm2.Hide
End Sub

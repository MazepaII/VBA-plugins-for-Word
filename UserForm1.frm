VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "��������"
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
If ������_����� = True Then
   Set objshell = Nothing
   Set ie = Nothing
   Set ������������ = Nothing
   ie.Quit
End If
Stop
Application.ScreenUpdating = False '��������� ���������� ���������
End Sub

Sub UserForm_Activate()
'On Error GoTo cancelhandler
'Application.EnableCancelKey = xlerrorhandler
'Application.EnableCancelKey = xlDisabled
If �����_��������_�_������_����_������ = True Then ��������_�_������_����_������
If �����_������_���_�_���� = True Then ������_���_�_����
If �����_���������_������������_�_������_�_���������_�_������� = True Then ���������_������������_�_������_�_���������_�_�������
If �����_�������_������� = True Then �������_�������
If �����_������_�����_��_������� = True Then ������_�����_��_�������
If �����_������_�_�����������_������_���������_��������� = True Then ������_�_�����������_������_���������_���������
If �����_��������_����_������� = True Then ��������_����_�������
If �����_��������_�����_������ = True Then ��������_�����_������

'cancelhandler:
'Application.EnableCancelKey = xlInterrupt
'If Err.Number = 18 Then MsgBox "5"
'Application.ScreenUpdating = True '�������� ���������� ���������1.Hide
'Application.EnableCancelKey = xlInterrupt
'UserForm1.Repaint

Unload UserForm1 '������ ����� � ��� ������ �� ���
'UserForm1.Hide  '������ ����� ������ �� ��� �� ���������
End Sub


'Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'Stop
'End Sub

Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii = 27 Then Stop 'esc
Application.ScreenUpdating = False '��������� ���������� ���������
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = 0 Then Stop '���� ������� ������
Application.ScreenUpdating = False '��������� ���������� ���������
End Sub

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "����� ���������!"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3570
   OleObjectBlob   =   "UserForm1.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()

End Sub

Public Sub Label1_Click()

End Sub

Private Sub UserForm_Activate()

UserForm1.Height = Label1.Top + Label1.Height + 20 + 15 + 15 ''������ ������� �������� + ������ ������� �������� + 10 ������� ����� ������ ��������� + 15 �� ��������� + 15 �� �����
UserForm1.Width = Label1.Left + Label1.Width + 15
''CommandButton1.Top = Label1.Height + 20
''CommandButton1.Left = UserForm1.Width / 2 - CommandButton1.Width / 2
Label2.Top = Label1.Height + 20
Label2.Left = UserForm1.Width / 2 - Label2.Width / 2
End Sub

Private Sub UserForm_Click()

End Sub

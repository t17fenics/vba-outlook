VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''��������� ����������� ��� ������� outlook
Private Sub Application_NewMail()

''�������� � 3 ������� ����� �������� �������
SleepVB ("3")

''������ ������ ��������� �������
RulesRun

End Sub

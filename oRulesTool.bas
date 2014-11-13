Attribute VB_Name = "oRulesTool"
''���������� ���������� message, ���������� ���������� � ������������ �������
Public message As String

''��������� ����������� ������� � outlook
Public Sub RulesRun()

''���������� ����������
Dim MyRules As Rules
Dim check As Byte

''��������� ����������
Set MyRules = Application.Session.DefaultStore.GetRules()
Rule = 1

''���� �������� ������
For Each Rule In MyRules
    
    ''����� �����������, ����� ����� ������������ ������� �� �������� ����� �� ������ ����������
    If Rule <= MyRules.Count Then
        
        ''���������� �������
        MyRules.Item(Rule).Execute
        
        ''������� ������ ������������� �� 1
        Rule = Rule + 1
DoEvents
     End If

Next

''����� ��������������� ��������� � ������� �������� � �������������� ��������
UFcontent

End Sub

''��������� ����������� ���������� ����� UserForm1
Public Sub UFcontent()

''���������� ����������
Dim oNamespace As Outlook.NameSpace
Dim oChildFolder As Outlook.MAPIFolder

''��������� ����������
Set oNamespace = Application.GetNamespace("MAPI")
Set oMailFolder = oNamespace.Folders("nikolai.karpov@heineken.com")
message = ""

''�������� ������������� �����
CheckNM oMailFolder

''���������� Label1 � UserForm1 ����������� � ������������� �������
UserForm1.Label1.Caption = message

''�������� � ��������� �����
UserForm1.Show

''�������� ������ �� �������� �����, � ������� ��������� ��������
oTimer (10)

''�������� �����
Unload UserForm1

End Sub

''��������� ������ ������������� �����
Public Sub CheckNM(ByVal ofolder As MAPIFolder)

''���������� ����������
Dim oChildFolder As Outlook.MAPIFolder

''���� ������������ ��� ����� � �������� �����
For Each oChildFolder In ofolder.Folders

    ''��� ������ �� ����� ���������� �������� ������� ������������� �����
    If oChildFolder.UnReadItemCount <> 0 Then
        ''���� ������������� ������ ������� - ���������� ���������� � ��� � ���������� message
        message = message + oChildFolder & " - " & oChildFolder.UnReadItemCount & vbCrLf
    End If
    DoEvents
    ''������� ��������� CheckNM ��� ������ �� ��������
    CheckNM oChildFolder
    
Next

End Sub

''��������� ��������� ������ ��������� ������� ������� � Seconds, �� ������ ������ � Label2
Public Sub oTimer(Seconds)

''���� ���������� �������
Do While Seconds > 0
    ''������������� �������� ���� Caption � Label2
    UserForm1.Label2.Caption = Seconds
    
    ''���������������� �����
    UserForm1.Repaint
    
    ''VB ������������� �� 1 ���
    SleepVB (1)
    
    ''�������� Second ����������� �� �������
    Seconds = Seconds - 1

Loop

End Sub

''��������� ������������������ ���������� ������� �� ���������� ������ ��������� � Seconds
Sub SleepVB(Seconds)

''���������� ����������
Dim Start

''��������� Start ������ �������� ������� � ��������
Start = timer

''��������� ���� �� ��� ���, ���� ������� ����� �� ������ ������ ������������ + Second
Do While timer < Start + Seconds

''������� ���������� ��
DoEvents

Loop

End Sub

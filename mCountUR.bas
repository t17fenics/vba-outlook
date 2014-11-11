Attribute VB_Name = "mCountUR"
Public mess As String
Public Sub countUR()

Dim Start
Dim oNamespace As Outlook.NameSpace
Dim oChildFolder As Outlook.MAPIFolder
mess = "Прочтите новые сообщения!" & vbCrLf & vbCrLf
Set oNamespace = Application.GetNamespace("MAPI")
Set oChildFolder = oNamespace.Folders("nikolai.karpov@heineken.com")
subfolder oChildFolder


UserForm1.Label1.Caption = mess
UserForm1.Show
''UserForm1.Hide
SleepVB ("3")

Unload UserForm1
''MsgBox (mess)



mess = ""
End
End Sub

Public Sub subfolder(ByVal ofolder As MAPIFolder)
Dim oChildFolder As Outlook.MAPIFolder

For Each oChildFolder In ofolder.Folders

    If oChildFolder.UnReadItemCount <> 0 Then
        mess = mess + oChildFolder & " - " & oChildFolder.UnReadItemCount & vbCrLf
        directory = oChildFolder
''        Debug.Print oChildFolder
        
        
    End If
    subfolder oChildFolder
    
Next

End Sub

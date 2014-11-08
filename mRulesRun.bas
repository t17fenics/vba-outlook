Attribute VB_Name = "mRulesRun"
Public Sub RulesRun()

Dim MyRules As Rules
Dim corul As Byte
Dim check As Byte
Dim myfold As Folders




Set MyRules = Application.Session.DefaultStore.GetRules()



corul = MyRules.Count
check = 1

For Each Rule In MyRules

    If corul >= check Then

        MyRules.Item(check).Execute
''        MsgBox ("ѕравило " & Chr(34) & MyRules.Item(check).Name & Chr(34) & " выполено.")
''        Debug.Print ("ѕравило " & Chr(34) & MyRules.Item(check).name & Chr(34) & " выполено.")
        
        check = check + 1

     End If

Next

countUR



End Sub


Sub SleepVB(Seconds)
' ожидание Seconds секунд
Dim Start
Start = Timer ' текущее врем€ в секундах
Do While Timer < Start + Seconds
' обеспечивает параллельное выполнение других процессов
DoEvents
Loop
End Sub



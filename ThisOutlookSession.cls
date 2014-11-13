VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''процедура выполняется при запуске outlook
Private Sub Application_NewMail()

''Задержка в 3 секунды перед запуском макроса
SleepVB ("3")

''Запуск первой процедуры макроса
RulesRun

End Sub

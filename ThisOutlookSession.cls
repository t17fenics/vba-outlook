VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)

End Sub

Private Sub Application_NewMail()

SleepVB ("3")
RulesRun

End Sub

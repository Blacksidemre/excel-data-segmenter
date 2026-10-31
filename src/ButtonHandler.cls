VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ButtonHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents Button As MSForms.CommandButton
Attribute Button.VB_VarHelpID = -1
Public Owner As Object

Private Sub Button_Click()
    On Error GoTo Failed
    Owner.HandleAction CStr(Button.Tag)
    Exit Sub
Failed:
    LogEvent "Arayüz", Err.Number, Err.Description
    MsgBox Err.Description, vbExclamation, "ExcelOps"
End Sub

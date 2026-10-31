Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    ExcelOpsAc
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    If Busy Then
        Cancel = True
        MsgBox "Önce işlemi durdurun veya tamamlanmasını bekleyin.", vbInformation
    End If
End Sub

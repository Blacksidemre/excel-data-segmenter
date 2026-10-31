VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSettings
   Caption = "ExcelOps Ayarları"
   ClientHeight = 9000
   ClientWidth = 11000
   StartUpPosition = 1
End
Attribute VB_Name = "frmSettings"
Attribute VB_Base = "0{C8C820B0-8D8B-40AE-B31D-1EC2709637F4}{1F1F53A9-164F-435D-BA89-492FAFD0A3AD}"
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private handlers As Collection
Private fieldCount As Long

Private Sub UserForm_Initialize()
    Dim ws As Worksheet, r As Long, c As Object, y As Single, key As String, choices As String, v As Variant
    Set handlers = New Collection: Set ws = ThisWorkbook.Worksheets("Ayarlar")
    StyleForm Me, "ExcelOps | Ayarlar", 760, 620
    AddLabel Me, "title", "Ayarlar · değerleri değiştirip Uygula düğmesine basın", 18, 12, 710, 26
    AddButton Me, handlers, "applyTop", "Uygula", "apply", 18, 46, 150
    AddButton Me, handlers, "second", "İkinci dosya seç", "second", 184, 46, 180
    AddButton Me, handlers, "dismissTop", "Vazgeç", "close", 380, 46, 150
    fieldCount = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To fieldCount
        y = 92 + (r - 2) * 44: key = CStr(ws.Cells(r, 1).Value2)
        AddLabel Me, "label" & r, CStr(ws.Cells(r, 3).Value2), 18, y, 230, 36
        choices = CStr(ws.Cells(r, 5).Value2)
        If choices <> "" Then
            Set c = Me.Controls.Add("Forms.ComboBox.1", "field" & r, True)
            c.Style = 2
            For Each v In Split(choices, ";"): c.AddItem CStr(v): Next
        Else
            Set c = Me.Controls.Add("Forms.TextBox.1", "field" & r, True)
        End If
        c.Left = 254: c.Top = y: c.Width = 450: c.Height = 22
        c.Tag = key: c.Value = CStr(ws.Cells(r, 2).Value2)
        c.ControlTipText = CStr(ws.Cells(r, 4).Value2)
    Next
    Me.ScrollBars = 2: Me.ScrollHeight = y + 68
    AddButton Me, handlers, "applyBottom", "Uygula", "apply", 254, y + 36, 180
End Sub

Public Sub HandleAction(ByVal action As String)
    Dim r As Long, c As Object, path As String
    Select Case action
        Case "apply"
            For r = 2 To fieldCount
                Set c = Me.Controls("field" & r): PutConfig CStr(c.Tag), CStr(c.Value)
            Next
            If Not ThisWorkbook.ReadOnly Then ThisWorkbook.Save
            Set handlers = Nothing
            Unload Me
        Case "second"
            path = PickFile()
            If path <> "" Then
                For r = 2 To fieldCount
                    Set c = Me.Controls("field" & r)
                    If c.Tag = "IkinciDosya" Then c.Value = path
                Next
            End If
        Case "close"
            Set handlers = Nothing
            Unload Me
    End Select
End Sub

Private Sub UserForm_Terminate()
    Set handlers = Nothing
End Sub

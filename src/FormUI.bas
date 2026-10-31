Attribute VB_Name = "FormUI"
Option Explicit

Public Sub StyleForm(ByVal form As Object, ByVal title As String, ByVal width As Single, ByVal height As Single)
    Dim control As Object
    ' The persisted designer is intentionally empty; controls are built at runtime.
    For Each control In form.Controls: control.Visible = False: Next
    form.Caption = title: form.Width = width: form.Height = height
    form.BackColor = RGB(244, 247, 251)
    form.Font.Name = "Segoe UI": form.Font.Size = 10
End Sub

Public Sub AddLabel(ByVal form As Object, ByVal name As String, ByVal caption As String, ByVal x As Single, ByVal y As Single, ByVal width As Single, Optional ByVal height As Single = 20)
    Dim c As Object
    Set c = form.Controls.Add("Forms.Label.1", name, True)
    c.Caption = caption: c.Left = x: c.Top = y: c.Width = width: c.Height = height
    c.BackStyle = 0: c.ForeColor = RGB(24, 49, 70)
End Sub

Public Sub AddButton(ByVal form As Object, ByVal handlers As Collection, ByVal name As String, ByVal caption As String, ByVal action As String, ByVal x As Single, ByVal y As Single, ByVal width As Single)
    Dim c As MSForms.CommandButton, handler As ButtonHandler
    Set c = form.Controls.Add("Forms.CommandButton.1", name, True)
    c.Caption = caption: c.Tag = action: c.Left = x: c.Top = y: c.Width = width: c.Height = 30
    c.BackColor = RGB(24, 49, 70): c.ForeColor = vbWhite: c.TakeFocusOnClick = False
    Set handler = New ButtonHandler: Set handler.Button = c: Set handler.Owner = form
    handlers.Add handler
End Sub

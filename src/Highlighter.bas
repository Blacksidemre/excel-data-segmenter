Attribute VB_Name = "Highlighter"
Option Explicit

Public Sub ColorRows(ByVal ws As Worksheet)
    Dim groups As Object, a As Variant, r As Long, first As Long, last As Long, col As Long
    Dim k As String, previous As String, start As Long, color As Long, maxCol As Long
    first = HeaderRow() + 1: last = LastRow(ws): maxCol = LastCol(ws)
    col = ColumnNumber(ws, Config("GrupSutunu")): a = ReadColumn(ws, first, last, col)
    Set groups = NewDict(): start = first
    For r = 1 To UBound(a, 1) + 1
        If r <= UBound(a, 1) Then k = KeyText(a(r, 1)) Else k = vbNullChar
        If r = 1 Then previous = k
        If k <> previous Then
            If previous <> "" Then
                If Not groups.Exists(previous) Then groups.Add previous, GroupColor(groups.Count, Config("Palet"))
                color = CLng(groups(previous))
                ws.Range(ws.Cells(start, 1), ws.Cells(first + r - 2, maxCol)).Interior.Color = color
            End If
            start = first + r - 1: previous = k
        End If
        If r Mod 2000 = 0 Then Pulse "Renklendirme: " & r - 1
    Next
    LogEvent "Renklendir", 0, groups.Count & " farklı grup rengi uygulandı."
End Sub

Public Function GroupColor(ByVal index As Long, ByVal palette As String) As Long
    Dim v As Double, red As Long, green As Long, blue As Long, base As Long, span As Long
    If palette = "Canli" Then
        base = 64: span = 192
    Else
        base = 128: span = 128
    End If
    ' Odd multiplier is bijective modulo a power of two: no repeated RGB in capacity.
    If index >= CDbl(span) ^ 3 Then Fail "Paletin benzersiz renk kapasitesi aşıldı."
    v = CDbl(index) * 65521#
    v = v - Fix(v / (CDbl(span) ^ 3)) * (CDbl(span) ^ 3)
    red = CLng(v) Mod span: v = Fix(v / span)
    green = CLng(v) Mod span: blue = Fix(v / span)
    GroupColor = RGB(base + red, base + green, base + blue)
End Function

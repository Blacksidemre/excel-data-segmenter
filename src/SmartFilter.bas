Attribute VB_Name = "SmartFilter"
Option Explicit

Private Function ISODate(ByVal text As String) As Double
    Dim parts As Variant, value As Date
    parts = Split(text, "-")
    If UBound(parts) <> 2 Then Fail "Tarih filtresi yyyy-mm-dd biçiminde olmalı."
    On Error GoTo Invalid
    value = DateSerial(CInt(parts(0)), CInt(parts(1)), CInt(parts(2)))
    If Format$(value, "yyyy-mm-dd") <> text Then GoTo Invalid
    ISODate = CDbl(value): Exit Function
Invalid:
    Fail "Geçersiz filtre tarihi: " & text
End Function

Public Sub FilterRows(ByVal ws As Worksheet)
    Dim h As Long, last As Long, helper As Long, j As Long, col As Long, r As Long
    Dim a As Variant, mask() As Variant, mode As String, spec As String, parts As Variant
    Dim allowed As Object, v As Variant, k As String, keep As Boolean, low As Double, high As Double, d As Double
    Dim rules As Long
    h = HeaderRow(): last = LastRow(ws): helper = LastCol(ws) + 1
    If helper > ws.Columns.Count Then Fail "Filtre yardımcı sütunu için yer yok."
    ReDim mask(1 To last - h, 1 To 1)
    For r = 1 To UBound(mask, 1): mask(r, 1) = 1: Next
    For j = 1 To 3
        col = ColumnNumber(ws, Config("Filtre" & j & "Sutun"), True)
        If col > 0 Then
            rules = rules + 1
            If col >= helper Then Fail "Filtre sütunu veri aralığı dışında."
            mode = Config("Filtre" & j & "Tur"): spec = Config("Filtre" & j & "Deger")
            Set allowed = NewDict(): parts = Split(spec, ";")
            Select Case mode
                Case "Dahil", "Haric"
                    For Each v In parts: allowed(Trim$(CStr(v))) = True: Next
                Case "Tutar", "Tarih"
                    If UBound(parts) <> 1 Then Fail "Aralık iki sınır içermeli: alt;üst"
                    If mode = "Tarih" Then
                        low = ISODate(Trim$(parts(0))): high = ISODate(Trim$(parts(1)))
                        If low > high Then Fail "Tarih alt sınırı üst sınırdan büyük."
                        high = high + 1
                    Else
                        low = AsAmount(parts(0), 0): high = AsAmount(parts(1), 0)
                    End If
                    If low > high Then Fail "Filtre alt sınırı üst sınırdan büyük."
                Case Else: Fail "Filtre türü: Dahil, Haric, Tutar veya Tarih olmalıdır."
            End Select
            a = ReadColumn(ws, h + 1, last, col)
            For r = 1 To UBound(a, 1)
                keep = False
                If Not IsError(a(r, 1)) Then
                    If mode = "Dahil" Or mode = "Haric" Then
                        keep = allowed.Exists(KeyText(a(r, 1)))
                        If mode = "Haric" Then keep = Not keep
                    ElseIf IsNumeric(a(r, 1)) And Not IsEmpty(a(r, 1)) Then
                        d = CDbl(a(r, 1))
                        If mode = "Tarih" Then
                            If ws.Parent.Date1904 Then d = d + 1462
                            keep = d >= low And d < high
                        Else
                            keep = d >= low And d <= high
                        End If
                    End If
                End If
                If Not keep Then mask(r, 1) = 0
            Next
        End If
    Next
    If rules = 0 Then Fail "En az bir filtre sütunu tanımlayın."
    If ws.ListObjects.Count > 0 Then Fail "Akıllı filtre için tabloyu kaynakta normal aralığa dönüştürün; kaynak otomatik değiştirilmez."
    If ws.FilterMode Then ws.ShowAllData
    ws.AutoFilterMode = False
    ws.Rows(CStr(h + 1) & ":" & CStr(last)).Hidden = False
    ws.Cells(h, helper).Value2 = "ExcelOps_Filtre"
    ws.Cells(h + 1, helper).Resize(UBound(mask, 1), 1).Value2 = mask
    ws.Range(ws.Cells(h, 1), ws.Cells(last, helper)).AutoFilter Field:=helper, Criteria1:="1"
    ws.Columns(helper).Hidden = True
End Sub

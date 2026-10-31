Attribute VB_Name = "DataQuality"
Option Explicit

Private Sub Issue(ByRef message As String, ByVal text As String)
    If Len(message) > 0 Then message = message & "; "
    message = message & text
End Sub

Public Sub CheckQuality(ByVal ws As Worksheet)
    Dim data As Variant, counts As Object, allowed As Object, required As Variant, cols() As Long
    Dim keyCol As Long, amountCol As Long, dateCol As Long, expectedCol As Long, h As Long, last As Long, maxCol As Long
    Dim r As Long, c As Long, i As Long, k As String, message As String, v As Variant, d As Double
    Dim report As Worksheet, out() As Variant, issues As Long, n As Long, label As String
    h = HeaderRow(): last = LastRow(ws): maxCol = LastCol(ws)
    keyCol = ColumnNumber(ws, Config("AnahtarSutunu"))
    amountCol = ColumnNumber(ws, Config("TutarSutunu"), True)
    dateCol = ColumnNumber(ws, Config("TarihSutunu"), True)
    expectedCol = ColumnNumber(ws, Config("BeklenenSutun"), True)
    maxCol = Application.Max(maxCol, keyCol, amountCol, dateCol, expectedCol, 2)
    required = Split(Config("ZorunluSutunlar"), ";"): ReDim cols(LBound(required) To UBound(required))
    For i = LBound(required) To UBound(required)
        cols(i) = ColumnNumber(ws, CStr(required(i)), True)
        If cols(i) > maxCol Then maxCol = cols(i)
    Next
    data = ws.Range(ws.Cells(h + 1, 1), ws.Cells(last, maxCol)).Value2
    Set counts = NewDict(): Set allowed = NewDict()
    For Each v In Split(Config("BeklenenDegerler"), ";")
        k = Trim$(CStr(v))
        If k <> "" Then allowed(k) = True
    Next
    If expectedCol > 0 And allowed.Count = 0 Then Fail "Beklenen değerler listesi boş."
    For r = 1 To UBound(data, 1)
        If Not IsError(data(r, keyCol)) Then
            k = KeyText(data(r, keyCol))
            If k <> "" Then
                If Not counts.Exists(k) Then counts.Add k, 0&
                counts(k) = counts(k) + 1
            End If
        End If
    Next
    ReDim out(1 To UBound(data, 1), 1 To 3)
    For r = 1 To UBound(data, 1)
        message = "": k = ""
        If IsError(data(r, keyCol)) Then
            Issue message, "Anahtar Excel hatası"
        Else
            k = KeyText(data(r, keyCol))
            If k = "" Then
                Issue message, "Anahtar boş"
            ElseIf counts(k) > 1 Then
                Issue message, "Mükerrer anahtar (" & counts(k) & ")"
            End If
        End If
        For i = LBound(cols) To UBound(cols)
            c = cols(i)
            If c > 0 Then
                If IsError(data(r, c)) Then
                    Issue message, "Excel hatası: sütun " & c
                ElseIf KeyText(data(r, c)) = "" Then
                    Issue message, "Zorunlu alan boş: sütun " & c
                End If
            End If
        Next
        If amountCol > 0 Then
            v = data(r, amountCol)
            If IsError(v) Then
                Issue message, "Tutar Excel hatası"
            ElseIf IsEmpty(v) Or Not IsNumeric(v) Then
                Issue message, "Tutar boş/sayısal değil"
            ElseIf CDbl(v) < 0 Then
                Issue message, "Negatif tutar"
            End If
        End If
        If dateCol > 0 Then
            v = data(r, dateCol)
            If IsError(v) Then
                Issue message, "Tarih Excel hatası"
            ElseIf IsEmpty(v) Or Not IsNumeric(v) Then
                Issue message, "Tarih Excel seri değeri değil"
            Else
                d = CDbl(v)
                If ws.Parent.Date1904 Then d = d + 1462
                If d < 1 Or d >= 2958466# Then Issue message, "Tarih aralık dışında"
            End If
        End If
        If expectedCol > 0 Then
            If IsError(data(r, expectedCol)) Then
                Issue message, "Beklenen alan Excel hatası"
            ElseIf Not allowed.Exists(KeyText(data(r, expectedCol))) Then
                Issue message, "Beklenmeyen değer"
            End If
        End If
        If message <> "" Then
            issues = issues + 1
            out(issues, 1) = r + h: out(issues, 2) = k: out(issues, 3) = message
            ws.Range(ws.Cells(r + h, 1), ws.Cells(r + h, maxCol)).Interior.Color = RGB(255, 210, 210)
        End If
        If r Mod 2000 = 0 Then Pulse "Veri kontrolü: " & r
    Next
    label = "ExcelOps_Kalite"
    Do
        Set report = Nothing
        On Error Resume Next
        Set report = ws.Parent.Worksheets(label)
        On Error GoTo 0
        If report Is Nothing Then Exit Do
        n = n + 1: label = "ExcelOps_Kalite_" & n
    Loop
    Set report = ws.Parent.Worksheets.Add(After:=ws.Parent.Sheets(ws.Parent.Sheets.Count))
    report.Name = label: SetHeaders report, Array("Kaynak satırı", "Anahtar", "Kontrol sonucu")
    report.Columns("B:C").NumberFormat = "@"
    ' Assign only populated rows without duplicating the entire dataset.
    If issues > 0 Then report.Range("A2").Resize(issues, 3).Value2 = out
    If issues = 0 Then report.Cells(2, 3) = "Seçilen kurallara göre hata bulunamadı."
    FinishReport report
    LogEvent "Veri Kontrolü", 0, issues & " hatalı satır / " & UBound(data, 1)
End Sub

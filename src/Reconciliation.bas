Attribute VB_Name = "Reconciliation"
Option Explicit

Private Function IndexFile(ByVal wb As Workbook, ByVal ws As Worksheet, ByVal h As Long, ByVal keyCol As Long, ByVal amountCol As Long) As Object
    Dim d As Object, keys As Variant, amounts As Variant, k As String, v As Variant, r As Long
    Set d = NewDict()
    keys = ReadColumn(ws, h + 1, LastRow(ws), keyCol): amounts = ReadColumn(ws, h + 1, LastRow(ws), amountCol)
    For r = 1 To UBound(keys, 1)
        k = KeyText(keys(r, 1))
        If k = "" Then Fail wb.Name & " içinde boş mutabakat anahtarı; satır " & r + h
        If Not d.Exists(k) Then d.Add k, Array(0&, 0#, r + h)
        v = d(k): v(0) = v(0) + 1: v(1) = v(1) + AsAmount(amounts(r, 1), r + h): d(k) = v
        If r Mod 5000 = 0 Then Pulse "Mutabakat dizini: " & wb.Name & " / " & r
    Next
    Set IndexFile = d
End Function

Private Function LooseKey(ByVal key As String) As String
    Dim ch As Variant
    For Each ch In Array(" ", "-", ".", "_", "/"): key = Replace(key, CStr(ch), ""): Next
    LooseKey = key
End Function

Public Sub Reconcile(ByVal source As Workbook, ByVal ws As Worksheet, ByVal target As Workbook)
    Dim second As Workbook, ws2 As Worksheet, left As Object, right As Object, all As Object, possible As Object
    Dim k As Variant, norm As String, a As Variant, b As Variant, output() As Variant
    Dim i As Long, status As String, candidate As String, tolerance As Double, code As Long, description As String
    Dim report As Worksheet, summary As Worksheet, countA As Long, countB As Long, totalA As Double, totalB As Double
    On Error GoTo Failed
    tolerance = AsAmount(Config("Tolerans"), 0)
    If tolerance < 0 Then Fail "Tolerans negatif olamaz."
    Set second = OpenSource(Config("IkinciDosya"))
    Set ws2 = DataSheet(second, Config("IkinciSayfa"), HeaderRow(True))
    Set left = IndexFile(source, ws, HeaderRow(), ColumnNumber(ws, Config("AnahtarSutunu")), ColumnNumber(ws, Config("TutarSutunu")))
    Set right = IndexFile(second, ws2, HeaderRow(True), ColumnNumber(ws2, Config("IkinciAnahtar")), ColumnNumber(ws2, Config("IkinciTutar")))
    second.Close False: Set second = Nothing
    Set all = NewDict(): Set possible = NewDict()
    For Each k In left.Keys: all.Add k, True: Next
    For Each k In right.Keys
        If Not all.Exists(k) Then all.Add k, True
        norm = LooseKey(CStr(k))
        If Len(norm) > 0 Then
            If possible.Exists(norm) Then possible(norm) = "[Birden fazla aday]" Else possible.Add norm, CStr(k)
        End If
    Next
    If all.Count > target.Worksheets(1).Rows.Count - 1 Then Fail "Mutabakat sonuçları sayfa sınırını aşıyor."
    ReDim output(1 To all.Count, 1 To 10)
    For Each k In all.Keys
        i = i + 1: a = Array(0&, 0#, 0&): b = Array(0&, 0#, 0&): status = "Eşleşti": candidate = ""
        If left.Exists(k) Then a = left(k)
        If right.Exists(k) Then b = right(k)
        If Not right.Exists(k) Then
            status = "İkinci dosyada eksik": norm = LooseKey(CStr(k))
            If possible.Exists(norm) Then candidate = CStr(possible(norm))
        ElseIf Not left.Exists(k) Then
            status = "İkinci dosyada fazla"
        ElseIf Abs(a(1) - b(1)) > tolerance Or a(0) <> b(0) Then
            status = "Fark var"
        End If
        If a(0) > 1 Or b(0) > 1 Then status = status & "; mükerrer anahtar"
        output(i, 1) = CStr(k): output(i, 2) = a(0): output(i, 3) = b(0)
        output(i, 4) = a(0) - b(0): output(i, 5) = a(1): output(i, 6) = b(1)
        output(i, 7) = a(1) - b(1): output(i, 8) = status: output(i, 9) = candidate
        output(i, 10) = "A:" & a(2) & " / B:" & b(2)
        countA = countA + a(0): countB = countB + b(0): totalA = totalA + a(1): totalB = totalB + b(1)
    Next
    Set report = target.Worksheets(1): report.Name = "Mutabakat"
    SetHeaders report, Array("Anahtar", "A adet", "B adet", "Adet farkı A-B", "A tutar", "B tutar", "Tutar farkı A-B", "Durum", "Olası eşleşme (öneri)", "İlk kaynak satırları")
    report.Columns(1).NumberFormat = "@": report.Columns("H:J").NumberFormat = "@"
    report.Range("A2").Resize(all.Count, 10).Value2 = output
    report.Columns("E:G").NumberFormat = "#,##0.00": FinishReport report
    Set summary = target.Worksheets.Add(After:=report): summary.Name = "Kontrol Toplamlari"
    SetHeaders summary, Array("Ölçüm", "A", "B", "Fark A-B")
    summary.Cells(2, 1) = "Kayıt sayısı": summary.Cells(2, 2) = countA: summary.Cells(2, 3) = countB
    summary.Cells(3, 1) = "Toplam tutar": summary.Cells(3, 2) = totalA: summary.Cells(3, 3) = totalB
    summary.Range("D2:D3").FormulaR1C1 = "=RC[-2]-RC[-1]"
    summary.Cells(5, 1) = "Mükerrer anahtarların tutarları toplanır; öneriler otomatik eşleşme değildir."
    summary.Range("B3:D3").NumberFormat = "#,##0.00": summary.Columns.AutoFit
    Exit Sub
Failed:
    code = Err.Number: description = Err.Description
    On Error Resume Next
    If Not second Is Nothing Then second.Close False
    On Error GoTo 0
    Err.Raise code, "Reconciliation", description
End Sub

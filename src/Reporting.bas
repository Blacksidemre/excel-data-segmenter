Attribute VB_Name = "Reporting"
Option Explicit

Public Sub Summarize(ByVal source As Workbook, ByVal ws As Worksheet, ByVal target As Workbook)
    Dim groups As Object, a As Variant, money As Variant, dates As Variant, item As Variant
    Dim key As Variant, k As String, r As Long, h As Long, last As Long, amountCol As Long, dateCol As Long
    Dim amount As Double, d As Double, output() As Variant, report As Worksheet, i As Long
    h = HeaderRow(): last = LastRow(ws): Set groups = NewDict()
    a = ReadColumn(ws, h + 1, last, ColumnNumber(ws, Config("GrupSutunu")))
    amountCol = ColumnNumber(ws, Config("TutarSutunu"))
    dateCol = ColumnNumber(ws, Config("TarihSutunu"), True)
    money = ReadColumn(ws, h + 1, last, amountCol)
    If dateCol > 0 Then dates = ReadColumn(ws, h + 1, last, dateCol)
    For r = 1 To UBound(a, 1)
        k = KeyText(a(r, 1))
        If k = "" Then Fail "Raporlama anahtarı boş. Satır: " & r + h
        amount = AsAmount(money(r, 1), r + h)
        d = 0
        If dateCol > 0 Then d = AsSerialDate(dates(r, 1), source, r + h)
        If Not groups.Exists(k) Then groups.Add k, Array(0&, 0#, amount, d, d)
        item = groups(k): item(0) = item(0) + 1: item(1) = item(1) + amount
        If amount > item(2) Then item(2) = amount
        If dateCol > 0 Then
            If d < item(3) Then item(3) = d
            If d > item(4) Then item(4) = d
        End If
        groups(k) = item
        If r Mod 5000 = 0 Then Pulse "Özet: " & r
    Next
    Set report = target.Worksheets(1): report.Name = "Ozet"
    SetHeaders report, Array("Grup", "Kayıt sayısı", "Toplam tutar", "Ortalama", "En büyük kayıt", "İlk kayıt tarihi", "Son kayıt tarihi")
    ReDim output(1 To groups.Count, 1 To 7)
    For Each key In groups.Keys
        i = i + 1: item = groups(key): output(i, 1) = CStr(key)
        output(i, 2) = item(0): output(i, 3) = item(1): output(i, 4) = item(1) / item(0): output(i, 5) = item(2)
        If dateCol > 0 Then output(i, 6) = item(3): output(i, 7) = item(4)
    Next
    report.Columns(1).NumberFormat = "@"
    report.Range("A2").Resize(groups.Count, 7).Value2 = output
    report.Columns("C:E").NumberFormat = "#,##0.00"
    report.Columns("F:G").NumberFormat = "dd.mm.yyyy"
    FinishReport report
End Sub

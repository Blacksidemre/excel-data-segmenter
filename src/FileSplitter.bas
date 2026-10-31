Attribute VB_Name = "FileSplitter"
Option Explicit

Public Sub SplitFiles(ByVal wb As Workbook, ByVal ws As Worksheet, ByVal folder As String)
    Dim groups As Object, names As Object, rows As Collection, keys As Variant
    Dim a As Variant, fileNames As Variant, k As String, nm As String, key As Variant
    Dim r As Long, col As Long, nameCol As Long, h As Long, last As Long, i As Long
    Dim path As String, cw As Worksheet, manifest As Worksheet, manifestBook As Workbook
    Dim code As Long, description As String
    On Error GoTo Failed
    h = HeaderRow(): last = LastRow(ws)
    col = ColumnNumber(ws, Config("GrupSutunu")): nameCol = ColumnNumber(ws, Config("AdSutunu"))
    a = ReadColumn(ws, h + 1, last, col): fileNames = ReadColumn(ws, h + 1, last, nameCol)
    Set groups = NewDict(): Set names = NewDict()
    For r = 1 To UBound(a, 1)
        k = KeyText(a(r, 1))
        If k <> "" Then
            nm = KeyText(fileNames(r, 1))
            If nm = "" Then Fail "Dosya adı alanı boş. Satır: " & r + h
            If Not groups.Exists(k) Then
                Set rows = New Collection: groups.Add k, rows: names.Add k, nm
            ElseIf StrComp(CStr(names(k)), nm, vbTextCompare) <> 0 Then
                Fail "Aynı grup için birden fazla dosya adı var: " & k
            End If
            Set rows = groups(k): rows.Add r + h
        End If
    Next
    If groups.Count = 0 Then Fail "Ayırma sütununda boş olmayan anahtar yok."
    If groups.Count > 100 Then
        If MsgBox(groups.Count & " tam çalışma kitabı kopyası üretilecek. Disk ve süre ihtiyacı yüksek olabilir. Devam?", _
            vbYesNo + vbQuestion, "ExcelOps") <> vbYes Then Err.Raise 18, , "İşlem iptal edildi."
    End If
    Set manifestBook = Workbooks.Add(xlWBATWorksheet): Set manifest = manifestBook.Worksheets(1)
    manifest.Name = "Dosya Listesi"
    SetHeaders manifest, Array("Grup", "Kayıt sayısı", "Dosya")
    manifest.Columns("A:A").NumberFormat = "@": manifest.Columns("C:C").NumberFormat = "@"
    keys = groups.Keys
    For Each key In keys
        i = i + 1: Pulse "Dosya ayırma: " & i & " / " & groups.Count & " - " & CStr(key)
        Set CurrentCopy = CloneBook(wb, folder, CurrentTemp)
        Set cw = CurrentCopy.Worksheets(ws.Name)
        Set rows = groups(key)
        DeleteOtherRows cw, h + 1, last, rows
        nm = Replace(Config("AdFormati"), "{ad}", CStr(names(key)))
        nm = Replace(nm, "{grup}", CStr(key))
        nm = Replace(nm, "{sira}", CStr(i))
        nm = Replace(nm, "{tarih}", Format$(Date, "yyyy-mm-dd"))
        nm = Config("OnEk") & nm & Config("SonEk")
        path = FreePath(folder, nm, OutputExt(wb))
        SaveResult CurrentCopy, path, Config("Bicim")
        CurrentCopy.Close SaveChanges:=False: Set CurrentCopy = Nothing
        RemoveTemp CurrentTemp: CurrentTemp = ""
        manifest.Cells(i + 1, 1).Value2 = CStr(key)
        manifest.Cells(i + 1, 2).Value2 = rows.Count
        manifest.Cells(i + 1, 3).Value2 = path
        LogEvent "Dosya Ayır", 0, rows.Count & " kayıt: " & path
    Next
    FinishReport manifest
    SaveResult manifestBook, FreePath(folder, "_Dosya_Listesi", ".xlsx"), "xlsx"
    manifestBook.Close False
    Exit Sub
Failed:
    code = Err.Number: description = Err.Description
    On Error Resume Next
    If Not manifestBook Is Nothing Then manifestBook.Close False
    On Error GoTo 0
    Err.Raise code, "FileSplitter", description
End Sub

Private Sub DeleteOtherRows(ByVal ws As Worksheet, ByVal first As Long, ByVal last As Long, ByVal keep As Collection)
    ' Sorted original row indices; remove complementary runs bottom-up, including hidden rows.
    Dim i As Long, bottom As Long, top As Long, areas As Long, pending As Range, block As Range
    bottom = last
    For i = keep.Count To 0 Step -1
        If i = 0 Then top = first Else top = CLng(keep(i)) + 1
        If bottom >= top Then
            Set block = ws.Rows(CStr(top) & ":" & CStr(bottom))
            If pending Is Nothing Then Set pending = block Else Set pending = Application.Union(pending, block)
            areas = areas + 1
            If areas = 128 Then
                pending.EntireRow.Delete: Set pending = Nothing: areas = 0
                Pulse "Kopyada satırlar ayrıştırılıyor..."
            End If
        End If
        bottom = top - 2
    Next
    If Not pending Is Nothing Then pending.EntireRow.Delete
End Sub

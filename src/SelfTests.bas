Attribute VB_Name = "SelfTests"
Option Explicit

Private Sub Check(ByVal condition As Boolean, ByVal label As String)
    If Not condition Then Err.Raise vbObjectError + 4000, "ExcelOpsSelfTest", "FAIL: " & label
    LogEvent "Kabul Testi", 0, "PASS: " & label
End Sub

Public Sub ExcelOpsSelfTest()
    Dim state As AppState, saved As Variant, cfg As Worksheet, source As Workbook, second As Workbook
    Dim report As Workbook, probe As Workbook, ws As Worksheet, hidden As Worksheet, rows As Variant
    Dim root As String, folder As String, bPath As String, temp As String, i As Long, code As Long, description As String
    If Busy Then Exit Sub
    Busy = True: CancelRequested = False
    Set cfg = ThisWorkbook.Worksheets("Ayarlar"): saved = cfg.UsedRange.Value2
    On Error GoTo Failed
    SuspendApp state
    root = Environ$("TEMP"): folder = RunFolder(root, "ExcelOps_KabulTesti")
    PutConfig "Baslik", "1": PutConfig "GrupSutunu", "M": PutConfig "AdSutunu", "M"
    PutConfig "AdFormati", "{ad}": PutConfig "OnEk", "": PutConfig "SonEk", "": PutConfig "Bicim", "xlsx"
    PutConfig "AnahtarSutunu", "A": PutConfig "TutarSutunu", "N": PutConfig "TarihSutunu", "O"
    PutConfig "Palet", "Pastel": PutConfig "ZorunluSutunlar", "A;M": PutConfig "BeklenenSutun", ""
    PutConfig "IkinciSayfa", "Veri": PutConfig "IkinciBaslik", "1": PutConfig "IkinciAnahtar", "A": PutConfig "IkinciTutar", "N"
    PutConfig "Tolerans", CStr(0.01)
    Set source = Workbooks.Add(xlWBATWorksheet): Set ws = source.Worksheets(1): ws.Name = "Veri"
    source.Date1904 = False
    For i = 1 To 15: ws.Cells(1, i) = "Alan" & i: Next
    rows = Array(Array("K1", " A ", 100), Array("K2", "A", 150), Array("K3", "B*", -20), _
        Array("K3", "B*", 80), Array("K5", "C?", 90), Array("K6", "C?", 200))
    For i = 0 To 5
        ws.Cells(i + 2, 1) = rows(i)(0): ws.Cells(i + 2, 13) = rows(i)(1)
        ws.Cells(i + 2, 14) = rows(i)(2): ws.Cells(i + 2, 15) = CDbl(DateSerial(2026, 10, 25 + i))
    Next
    ws.Columns(13).ColumnWidth = 23: ws.Columns(15).NumberFormat = "dd.mm.yyyy"
    ws.Range("A1:O7").AutoFilter Field:=13, Criteria1:="A"
    Set hidden = source.Worksheets.Add: hidden.Name = "Gizli": hidden.Cells(1, 1).Formula = "=1+1"
    hidden.Visible = xlSheetVeryHidden
    source.SaveAs Filename:=folder & "\Kaynak.xlsx", FileFormat:=xlOpenXMLWorkbook
    SplitFiles source, ws, folder
    Set probe = Workbooks.Open(folder & "\A.xlsx", UpdateLinks:=0, ReadOnly:=True)
    Check LastRow(probe.Worksheets("Veri")) = 3, "Ayırma: trim ile iki A kaydı"
    Check probe.Worksheets("Gizli").Visible = xlSheetVeryHidden, "Gizli sayfa korundu"
    Check probe.Worksheets("Gizli").Cells(1, 1).Formula = "=1+1", "Formül korundu"
    Check probe.Worksheets("Veri").Columns(13).ColumnWidth = 23, "Sütun genişliği korundu"
    Check probe.Worksheets("Veri").AutoFilterMode, "Filtre yapısı korundu"
    probe.Close False: Set probe = Nothing
    Set probe = Workbooks.Open(folder & "\B_.xlsx", UpdateLinks:=0, ReadOnly:=True)
    Check LastRow(probe.Worksheets("Veri")) = 3, "Yıldız gerçek karakter; gizli satırlar ayrıldı"
    Check probe.Worksheets("Veri").Cells(2, 13).Value2 = "B*", "B grubunda yalnız B kayıtları"
    probe.Close False: Set probe = Nothing
    Check LastRow(ws) = 7 And ws.Cells(2, 13).Value2 = " A ", "Kaynak satırlar değişmedi"
    Set report = Workbooks.Add(xlWBATWorksheet): Summarize source, ws, report
    Check report.Worksheets(1).Cells(2, 3).Value2 = 250, "Özet A toplamı"
    Check report.Worksheets(1).Cells(3, 3).Value2 = 60, "Özet B toplamı"
    report.Close False: Set report = Nothing
    Set probe = CloneBook(source, folder, temp): ColorRows probe.Worksheets("Veri")
    Check probe.Worksheets("Veri").Cells(2, 1).Interior.Color = probe.Worksheets("Veri").Cells(3, 1).Interior.Color, "Aynı grup aynı renk"
    Check probe.Worksheets("Veri").Cells(2, 1).Interior.Color <> probe.Worksheets("Veri").Cells(4, 1).Interior.Color, "Farklı grup farklı renk"
    CheckQuality probe.Worksheets("Veri")
    Check LastRow(probe.Worksheets("ExcelOps_Kalite")) = 3, "Mükerrer kayıtlar ve negatif tutar kontrolü"
    PutConfig "Filtre1Sutun", "N": PutConfig "Filtre1Tur", "Tutar": PutConfig "Filtre1Deger", "50;200"
    PutConfig "Filtre2Sutun", "": PutConfig "Filtre3Sutun", ""
    FilterRows probe.Worksheets("Veri")
    Check Application.WorksheetFunction.CountIf(probe.Worksheets("Veri").Range("P2:P7"), 1) = 5, "Tutar aralığı iki sınır dahil"
    probe.Close False: Set probe = Nothing: RemoveTemp temp: temp = ""
    bPath = folder & "\Ikinci.xlsx": source.SaveCopyAs bPath
    Set second = Workbooks.Open(bPath, UpdateLinks:=0): second.Worksheets("Veri").Cells(2, 14) = 110
    second.Close True: Set second = Nothing: PutConfig "IkinciDosya", bPath
    Set report = Workbooks.Add(xlWBATWorksheet): Reconcile source, ws, report
    Check report.Worksheets("Mutabakat").Cells(2, 7).Value2 = -10, "Mutabakat tutar farkı"
    Check InStr(report.Worksheets("Mutabakat").Cells(4, 8).Value2, "mükerrer") > 0, "Mükerrer mutabakat anahtarı"
    report.Close False: Set report = Nothing
    source.Close False: Set source = Nothing
    cfg.UsedRange.Value2 = saved
    RestoreApp state: Busy = False
    MsgBox "Kabul testleri geçti. Ayrıntılar Log sayfasında. Test dosyaları: " & folder, vbInformation
    Exit Sub
Failed:
    code = Err.Number: description = Err.Description
    On Error Resume Next
    If Not probe Is Nothing Then probe.Close False
    If Not CurrentCopy Is Nothing Then CurrentCopy.Close False
    Set CurrentCopy = Nothing
    If Not report Is Nothing Then report.Close False
    If Not source Is Nothing Then source.Close False
    If Not second Is Nothing Then second.Close False
    RemoveTemp CurrentTemp: CurrentTemp = "": RemoveTemp temp
    cfg.UsedRange.Value2 = saved
    On Error GoTo 0
    RestoreApp state: Busy = False
    LogEvent "Kabul Testi", code, description
    MsgBox description & vbCrLf & "Test dosyaları: " & folder, vbCritical
End Sub

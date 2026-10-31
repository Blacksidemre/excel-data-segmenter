Attribute VB_Name = "ExcelOps"
Option Explicit
Public CurrentCopy As Workbook
Public CurrentTemp As String

Public Sub ExcelOpsAc()
    If Busy Then Exit Sub
    frmMain.Show vbModeless
End Sub

Public Sub RunOperation(ByVal action As String)
    Dim state As AppState, source As Workbook, ws As Worksheet, folder As String, path As String
    Dim code As Long, description As String, suspended As Boolean, preview As Boolean
    Dim reportSheet As Worksheet
    If Busy Then Exit Sub
    If Config("KaynakDosya") = "" Then
        MsgBox "Önce Kaynak Seç ve Ayarlar düğmelerini kullanın.", vbInformation: Exit Sub
    End If
    Busy = True: CancelRequested = False
    On Error GoTo Failed
    suspended = True: SuspendApp state
    Set source = OpenSource(Config("KaynakDosya"))
    Set ws = DataSheet(source, Config("Sayfa"), HeaderRow())
    If ws.ProtectContents Or source.ProtectStructure Then Fail "Korumalı sayfa/kitap desteklenmiyor. Korumasız bir kaynak kopyası kullanın."
    If IsNull(ws.Range(ws.Cells(HeaderRow() + 1, 1), ws.Cells(LastRow(ws), LastCol(ws))).MergeCells) Then Fail "Veri alanında birleşik hücreler var."
    If ws.Range(ws.Cells(HeaderRow() + 1, 1), ws.Cells(LastRow(ws), LastCol(ws))).MergeCells Then Fail "Veri alanında birleşik hücreler var."
    preview = action = "Renklendir" And Config("RenkDosya") = "Hayir"
    If Not preview Then folder = RunFolder(Config("Konum"), Config("Klasor"))
    Select Case action
        Case "Dosya Ayır"
            SplitFiles source, ws, folder
        Case "Renklendir", "Veri Kontrolü", "Akıllı Filtre"
            If preview Then
                ColorRows ws
            Else
                Set CurrentCopy = CloneBook(source, folder, CurrentTemp)
                Set ws = CurrentCopy.Worksheets(ws.Name)
                Select Case action
                    Case "Renklendir": ColorRows ws
                    Case "Veri Kontrolü": CheckQuality ws
                    Case "Akıllı Filtre": FilterRows ws
                End Select
                path = FreePath(folder, action, OutputExt(source))
                SaveResult CurrentCopy, path, Config("Bicim")
            End If
        Case "Özet Çıkar", "Mutabakat Yap"
            Set CurrentCopy = Workbooks.Add(xlWBATWorksheet)
            CurrentCopy.Date1904 = False
            If action = "Özet Çıkar" Then Summarize source, ws, CurrentCopy Else Reconcile source, ws, CurrentCopy
            For Each reportSheet In CurrentCopy.Worksheets: reportSheet.Calculate: Next
            path = FreePath(folder, action, ".xlsx")
            SaveResult CurrentCopy, path, "xlsx"
        Case Else: Fail "Bilinmeyen işlem."
    End Select
    If Not CurrentCopy Is Nothing Then CurrentCopy.Close False
    Set CurrentCopy = Nothing
    RemoveTemp CurrentTemp: CurrentTemp = ""
    If preview Then
        Set source = Nothing
        description = "Renkler salt okunur açık kaynakta gösteriliyor. Kaynak diskte değişmedi. Kalıcılaştırmak için Farklı Kaydet kullanın."
    Else
        source.Close False: Set source = Nothing
        description = "İşlem tamamlandı: " & folder
    End If
    LogEvent action, 0, description
    RestoreApp state: Busy = False
    MsgBox description, vbInformation, "ExcelOps": Exit Sub
Failed:
    code = Err.Number: description = Err.Description
    On Error Resume Next
    If Not CurrentCopy Is Nothing Then CurrentCopy.Close False
    Set CurrentCopy = Nothing
    If Not source Is Nothing Then source.Close False
    RemoveTemp CurrentTemp: CurrentTemp = ""
    On Error GoTo 0
    LogEvent action, code, description & " | Çıktı: " & folder
    If suspended Then RestoreApp state
    Busy = False
    MsgBox "İşlem tamamlanamadı (" & code & "): " & description & vbCrLf & _
        "Tamamlanan çıktılar varsa korunmuştur. Log sayfasını kontrol edin." & vbCrLf & folder, vbExclamation, "ExcelOps"
End Sub

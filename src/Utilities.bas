Attribute VB_Name = "Utilities"
Option Explicit

Public Type AppState
    Screen As Boolean
    Alerts As Boolean
    Events As Boolean
    Calculation As XlCalculation
    Security As Long
    Status As Variant
    CancelKey As XlEnableCancelKey
End Type
Public CancelRequested As Boolean
Public Busy As Boolean

Public Sub SuspendApp(ByRef state As AppState)
    With Application
        state.Screen = .ScreenUpdating: state.Alerts = .DisplayAlerts
        state.Events = .EnableEvents: state.Calculation = .Calculation
        state.Security = .AutomationSecurity: state.Status = .StatusBar
        state.CancelKey = .EnableCancelKey
        .ScreenUpdating = False: .DisplayAlerts = False: .EnableEvents = False
        .Calculation = xlCalculationManual: .AutomationSecurity = 3
        .EnableCancelKey = xlErrorHandler
    End With
End Sub

Public Sub RestoreApp(ByRef state As AppState)
    On Error Resume Next
    With Application
        .AutomationSecurity = state.Security: .Calculation = state.Calculation
        .EnableEvents = state.Events: .DisplayAlerts = state.Alerts
        .ScreenUpdating = state.Screen: .StatusBar = state.Status
        .EnableCancelKey = state.CancelKey
    End With
    On Error GoTo 0
End Sub

Public Sub Fail(ByVal message As String)
    Err.Raise vbObjectError + 2048, "ExcelOps", message
End Sub

Public Function NewDict() As Object
    Set NewDict = CreateObject("Scripting.Dictionary")
    NewDict.CompareMode = vbTextCompare
End Function

Public Function KeyText(ByVal value As Variant) As String
    If IsError(value) Then Fail "Anahtar alanında Excel hata değeri var. Önce Veri Kontrolü çalıştırın."
    If IsEmpty(value) Or IsNull(value) Then Exit Function
    KeyText = Trim$(Replace(CStr(value), ChrW(160), " "))
End Function

Public Function SafeName(ByVal value As String) As String
    Dim ch As Variant, i As Long, stem As String
    For Each ch In Array("\", "/", ":", "*", "?", Chr$(34), "<", ">", "|")
        value = Replace(value, CStr(ch), "_")
    Next
    For i = 0 To 31: value = Replace(value, Chr$(i), "_"): Next
    value = Trim$(value)
    Do While Len(value) > 0
        If Right$(value, 1) <> "." And Right$(value, 1) <> " " Then Exit Do
        value = Left$(value, Len(value) - 1)
    Loop
    If Len(value) = 0 Then value = "Adsiz"
    stem = UCase$(Split(value, ".")(0))
    If stem = "CON" Or stem = "PRN" Or stem = "AUX" Or stem = "NUL" Then value = "_" & value
    For i = 1 To 9
        If stem = "COM" & i Or stem = "LPT" & i Then value = "_" & value
    Next
    SafeName = Left$(value, 100)
End Function

Public Function LastRow(ByVal ws As Worksheet) As Long
    Dim c As Range
    Set c = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    If Not c Is Nothing Then LastRow = c.Row
End Function

Public Function LastCol(ByVal ws As Worksheet) As Long
    Dim c As Range
    Set c = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    If Not c Is Nothing Then LastCol = c.Column
End Function

Public Function ColumnNumber(ByVal ws As Worksheet, ByVal text As String, Optional ByVal optionalField As Boolean = False) As Long
    Dim n As Long
    text = UCase$(Trim$(text))
    If text = "" And optionalField Then Exit Function
    On Error GoTo Invalid
    If IsNumeric(text) Then
        If CDbl(text) <> Fix(CDbl(text)) Then GoTo Invalid
        n = CLng(text)
    Else
        If Len(text) < 1 Or Len(text) > 3 Or text Like "*[!A-Z]*" Then GoTo Invalid
        n = ws.Range(text & "1").Column
    End If
    If n < 1 Or n > ws.Columns.Count Then GoTo Invalid
    ColumnNumber = n
    Exit Function
Invalid:
    Fail "Geçersiz sütun: " & text & ". A, M, AA veya sütun numarası girin."
End Function

Public Function ReadColumn(ByVal ws As Worksheet, ByVal first As Long, ByVal last As Long, ByVal col As Long) As Variant
    Dim a() As Variant
    If last = first Then
        ReDim a(1 To 1, 1 To 1): a(1, 1) = ws.Cells(first, col).Value2: ReadColumn = a
    Else
        ReadColumn = ws.Range(ws.Cells(first, col), ws.Cells(last, col)).Value2
    End If
End Function

Public Function OpenSource(ByVal path As String) As Workbook
    Dim wb As Workbook, fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(path) Then Fail "Kaynak dosya bulunamadı: " & path
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, path, vbTextCompare) = 0 Then _
            Fail "Kaynak dosya zaten açık. Kaydedip kapatın; sonra tekrar seçin."
    Next
    Set OpenSource = Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=True, _
        IgnoreReadOnlyRecommended:=True, AddToMru:=False, Notify:=False)
End Function

Public Function DataSheet(ByVal wb As Workbook, ByVal name As String, ByVal header As Long) As Worksheet
    On Error Resume Next
    Set DataSheet = wb.Worksheets(name)
    On Error GoTo 0
    If DataSheet Is Nothing Then Fail "Sayfa bulunamadı: " & name
    If header < 1 Or header >= DataSheet.Rows.Count Then Fail "Başlık satırı geçersiz."
    If LastRow(DataSheet) <= header Then Fail "Başlığın altında veri bulunamadı."
End Function

Public Function NativeExt(ByVal wb As Workbook) As String
    Select Case wb.FileFormat
        Case xlOpenXMLWorkbook: NativeExt = ".xlsx"
        Case xlOpenXMLWorkbookMacroEnabled: NativeExt = ".xlsm"
        Case xlExcel12: NativeExt = ".xlsb"
        Case xlExcel8: NativeExt = ".xls"
        Case Else: Fail "Desteklenen kaynak biçimleri: xlsx, xlsm, xlsb, xls."
    End Select
End Function

Public Function RunFolder(ByVal base As String, ByVal label As String) As String
    Dim fso As Object, parent As String, path As String, n As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(base) Then Fail "Kayıt konumu bulunamadı. Var olan bir klasör seçin."
    parent = fso.BuildPath(base, SafeName(label))
    If Not fso.FolderExists(parent) Then fso.CreateFolder parent
    path = fso.BuildPath(parent, Format$(Now, "yyyymmdd_hhnnss"))
    Do While fso.FolderExists(path)
        n = n + 1: path = fso.BuildPath(parent, Format$(Now, "yyyymmdd_hhnnss") & "_" & n)
    Loop
    If Len(path) > 120 Then Fail "Kayıt yolu çok uzun. Daha kısa bir ana klasör seçin."
    fso.CreateFolder path: RunFolder = path
End Function

Public Function FreePath(ByVal folder As String, ByVal name As String, ByVal ext As String) As String
    Dim fso As Object, stem As String, p As String, i As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
    stem = SafeName(name): p = fso.BuildPath(folder, stem & ext)
    Do While fso.FileExists(p)
        i = i + 1: p = fso.BuildPath(folder, stem & "_" & i & ext)
    Loop
    If Len(p) > 218 Then Fail "Dosya yolu Excel sınırını aşıyor: " & p
    FreePath = p
End Function

Public Function CloneBook(ByVal source As Workbook, ByVal folder As String, ByRef tempPath As String) As Workbook
    tempPath = FreePath(folder, "_ExcelOps_gecici", NativeExt(source))
    source.SaveCopyAs tempPath
    Set CloneBook = Workbooks.Open(Filename:=tempPath, UpdateLinks:=0, ReadOnly:=False, AddToMru:=False)
End Function

Public Sub SaveResult(ByVal wb As Workbook, ByVal finalPath As String, ByVal format As String)
    If Dir$(finalPath) <> "" Then Fail "Çıktı zaten var; üzerine yazılmadı: " & finalPath
    If format = "xlsx" Then
        wb.CheckCompatibility = False
        wb.SaveAs Filename:=finalPath, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Else
        wb.SaveAs Filename:=finalPath, FileFormat:=wb.FileFormat, CreateBackup:=False
    End If
End Sub

Public Sub RemoveTemp(ByVal path As String)
    If Len(path) > 0 Then
        If Dir$(path) <> "" Then Kill path
    End If
End Sub

Public Sub Pulse(ByVal message As String)
    Application.StatusBar = message
    DoEvents
    If CancelRequested Then Err.Raise 18, "ExcelOps", "İşlem kullanıcı tarafından durduruldu."
End Sub

Public Sub LogEvent(ByVal action As String, ByVal code As Long, ByVal detail As String)
    Dim ws As Worksheet, r As Long
    On Error GoTo Unable
    Set ws = ThisWorkbook.Worksheets("Log")
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If r > ws.Rows.Count Then Exit Sub
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = action: ws.Cells(r, 3).Value = code
    ws.Cells(r, 4).NumberFormat = "@": ws.Cells(r, 4).Value2 = detail
    Exit Sub
Unable:
    Debug.Print "ExcelOps günlük yazılamadı: " & Err.Description
End Sub

Public Sub SetHeaders(ByVal ws As Worksheet, ByVal labels As Variant)
    Dim i As Long
    For i = LBound(labels) To UBound(labels): ws.Cells(1, i + 1).Value = labels(i): Next
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, UBound(labels) + 1))
        .Font.Bold = True: .Font.Color = vbWhite: .Interior.Color = RGB(24, 49, 70)
    End With
End Sub

Public Sub FinishReport(ByVal ws As Worksheet)
    ws.UsedRange.Columns.AutoFit
    ws.UsedRange.AutoFilter
End Sub

Public Function AsAmount(ByVal value As Variant, ByVal row As Long) As Double
    If IsError(value) Then Fail "Tutar alanında Excel hatası. Satır: " & row
    If IsEmpty(value) Or Trim$(CStr(value)) = "" Then Fail "Tutar boş. Satır: " & row
    If Not IsNumeric(value) Then Fail "Tutar sayısal değil. Satır: " & row
    AsAmount = CDbl(value)
End Function

Public Function AsSerialDate(ByVal value As Variant, ByVal wb As Workbook, ByVal row As Long) As Double
    Dim d As Double
    If IsError(value) Then Fail "Tarih alanında Excel hatası. Satır: " & row
    If Not IsNumeric(value) Or IsEmpty(value) Then Fail "Gerçek Excel tarihi gerekiyor. Satır: " & row
    d = CDbl(value)
    If wb.Date1904 Then d = d + 1462
    If d < 1 Or d >= 2958466# Then Fail "Geçersiz tarih. Satır: " & row
    AsSerialDate = d
End Function

Attribute VB_Name = "Settings"
Option Explicit

Public Function Config(ByVal key As String) As String
    Dim hit As Range
    Set hit = ThisWorkbook.Worksheets("Ayarlar").Columns(1).Find(What:=key, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False, SearchFormat:=False)
    If hit Is Nothing Then Fail "Ayar bulunamadı: " & key
    Config = CStr(hit.Offset(0, 1).Value2)
End Function

Public Sub PutConfig(ByVal key As String, ByVal value As String)
    Dim hit As Range
    Set hit = ThisWorkbook.Worksheets("Ayarlar").Columns(1).Find(What:=key, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False, SearchFormat:=False)
    If hit Is Nothing Then Fail "Ayar bulunamadı: " & key
    hit.Offset(0, 1).NumberFormat = "@": hit.Offset(0, 1).Value2 = value
End Sub

Public Function HeaderRow(Optional ByVal second As Boolean = False) As Long
    Dim v As String
    If second Then v = Config("IkinciBaslik") Else v = Config("Baslik")
    If Not IsNumeric(v) Then Fail "Başlık satırı tam sayı olmalıdır."
    If CDbl(v) < 1 Or CDbl(v) > 1048575 Or CDbl(v) <> Fix(CDbl(v)) Then Fail "Başlık satırı geçersiz."
    HeaderRow = CLng(v)
End Function

Public Function OutputExt(ByVal wb As Workbook) As String
    If Config("Bicim") = "xlsx" Then OutputExt = ".xlsx" Else OutputExt = NativeExt(wb)
End Function

Public Function PickFile() As String
    With Application.FileDialog(3)
        .Title = "İşlenecek Excel dosyasını seçin"
        .AllowMultiSelect = False: .Filters.Clear
        .Filters.Add "Excel", "*.xlsx;*.xlsm;*.xlsb;*.xls"
        If .Show = -1 Then PickFile = .SelectedItems(1)
    End With
End Function

Public Function PickFolder() As String
    With Application.FileDialog(4)
        .Title = "Çıktılar için ana klasörü seçin"
        If .Show = -1 Then PickFolder = .SelectedItems(1)
    End With
End Function

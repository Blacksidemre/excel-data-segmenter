VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain
   Caption = "ExcelOps"
   ClientHeight = 6600
   ClientWidth = 9300
   StartUpPosition = 1
End
Attribute VB_Name = "frmMain"
Attribute VB_Base = "0{77B36796-865B-409C-831C-2DC9BAF3D360}{5DC37A62-CB6C-4A60-9FF6-B0ED474E5387}"
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private handlers As Collection

Private Sub UserForm_Initialize()
    Dim labels As Variant, i As Long
    Set handlers = New Collection
    StyleForm Me, "ExcelOps | Excel Operasyon Platformu", 640, 460
    AddLabel Me, "title", "ExcelOps", 24, 18, 560, 32
    Me.Controls("title").Font.Size = 24: Me.Controls("title").Font.Bold = True
    AddLabel Me, "subtitle", "Kaynağı seçin, ayarları kontrol edin, işlemi başlatın.", 24, 58, 580
    AddLabel Me, "sourcePath", "", 24, 90, 580, 42
    AddButton Me, handlers, "choose", "Kaynak seç", "source", 24, 142, 180
    AddButton Me, handlers, "settings", "Ayarlar", "settings", 220, 142, 180
    AddButton Me, handlers, "folder", "Kayıt konumu", "folder", 416, 142, 180
    labels = Array("Dosya Ayır", "Renklendir", "Özet Çıkar", "Mutabakat Yap", "Veri Kontrolü", "Akıllı Filtre")
    For i = 0 To UBound(labels)
        AddButton Me, handlers, "run" & i, CStr(labels(i)), CStr(labels(i)), 24 + (i Mod 3) * 196, 194 + (i \ 3) * 48, 180
    Next
    AddLabel Me, "note", "Kaynak dosya diskte değiştirilmez. Çıktılar ayrı işlem klasörüne yazılır. Ayırmada diğer sayfalar tüm kayıtlarıyla korunur.", 24, 300, 570, 44
    AddButton Me, handlers, "stop", "İşlemi durdur", "cancel", 24, 363, 180
    AddButton Me, handlers, "help", "Kullanım kılavuzu", "help", 220, 363, 180
    AddButton Me, handlers, "close", "Menüyü kapat", "close", 416, 363, 180
    RefreshSource
End Sub

Private Sub RefreshSource()
    Me.Controls("sourcePath").Caption = "Kaynak: " & Config("KaynakDosya")
End Sub

Public Sub HandleAction(ByVal action As String)
    Dim path As String
    If action = "cancel" Then CancelRequested = True: Exit Sub
    If Busy Then Exit Sub
    Select Case action
        Case "source"
            path = PickFile()
            If path <> "" Then PutConfig "KaynakDosya", path
            RefreshSource
        Case "folder"
            path = PickFolder()
            If path <> "" Then PutConfig "Konum", path
        Case "settings": frmSettings.Show vbModal
        Case "help"
            MsgBox "Dosya içindeki Kılavuz sayfasını okuyun. Ayarları uyguladıktan sonra ilgili işlem düğmesini kullanın." & vbCrLf & _
                "Menüyü yeniden açmak: Alt+F8 > ExcelOpsAc.", vbInformation, "ExcelOps"
        Case "close": Me.Hide
        Case Else: RunOperation action
    End Select
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If Busy Then
        Cancel = True: CancelRequested = True
    Else
        Set handlers = Nothing
    End If
End Sub

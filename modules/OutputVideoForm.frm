VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OutputVideoForm 
   Caption         =   "导出视频"
   ClientHeight    =   4320
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6360
   OleObjectBlob   =   "OutputVideoForm.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "OutputVideoForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public vuse As String, oname As String, ofmt As String, otime As Single, oph As Single, ofps As Single, oq As Single
'Dim w As Single, h As Single, ow As Single, oh As Single
Private Const TIMES As String = " × "
Private originalFileName As String

Private Sub BrowerButton_Click()
    Dim isSave As Variant
    isSave = SaveAsFile(FileNameText.Text, FormatCombo.Text)
    If isSave <> False Then
        originalFileName = isSave
        SetFileNames originalFileName
    End If
End Sub

Private Sub CancelButton_Click()
    End
End Sub

Private Sub DurationSpin_SpinDown()
    Dim duration As Long
    duration = Val(DurationText.Value) \ 1 - 1
    If duration < 0 Then duration = 0
    DurationText.Value = duration
End Sub

Private Sub DurationSpin_SpinUp()
    DurationText.Value = Val(DurationText.Value) \ 1 + 1
End Sub

Private Sub OkButton_Click()
    SaveVideo _
        FileNameText.Text, _
        FormatCombo.Text, _
        UseNarrationsCheck.Value, _
        Val(DurationText.Text), _
        Val(OutputHeightCombo.Text), _
        Val(FpsCombo.Text), _
        Val(QualityText.Text)
    'OutputVideoForm.Hide
    End
End Sub

Private Sub QualitySpin_SpinDown()
    Dim quality As Long
    quality = Val(QualityText.Value) \ 1 - 1
    If quality < 0 Then quality = 0
    QualityText.Value = quality
End Sub

Private Sub QualitySpin_SpinUp()
    Dim quality As Long
    quality = Val(QualityText.Value) \ 1 + 1
    If quality > 100 Then quality = 100
    QualityText.Value = quality
End Sub

Private Sub ResetButton_Click()
    FormatCombo.List = Array(".mp4", ".wmv")
    OutputHeightCombo.List = Array("2160", "1080", "720", "480")
    FpsCombo.List = Array("120", "60", "30", "25", "15")
    DurationText.Text = 5
    FormatCombo.Text = ".mp4"
    SetFileNames originalFileName
    Dim width As Long, height As Long
    width = ActivePresentation.PageSetup.SlideWidth
    height = ActivePresentation.PageSetup.SlideHeight
    OrigDpiText.Caption = width & TIMES & height
    OutputWidthText.Caption = GetWidth(Val(OutputHeightCombo.Text)) & TIMES
    OutputHeightCombo.Text = 1080
    FpsCombo.Text = 60
    QualityText.Text = 100
    UseNarrationsCheck.Value = False
End Sub

Private Sub FpsCombo_Change()
    FpsCombo.Text = Val(FpsCombo.Text) \ 1
End Sub

Private Sub OutputHeightCombo_Change()
    OutputWidthText.Caption = GetWidth(Val(OutputHeightCombo.Text)) & TIMES
End Sub

Private Sub QualityText_Change()
    Dim quality As Long
    quality = Val(QualityText.Value) \ 1
    If quality < 0 Then quality = 0
    If quality > 100 Then quality = 100
    QualityText.Value = quality
End Sub

Private Sub UserForm_Initialize()
    Dim isSave As Variant
    isSave = SaveAsFile
    If isSave = False Then End
    originalFileName = isSave
    Call ResetButton_Click
End Sub

Private Function GetWidth(providedHeight As Long) As Long
    Dim width As Long, height As Long
    width = ActivePresentation.PageSetup.SlideWidth
    height = ActivePresentation.PageSetup.SlideHeight
    GetWidth = providedHeight / height * width \ 1
End Function

Private Function SaveAsFile(Optional fileName As String, Optional format As String) As Variant
    With Application.FileDialog(msoFileDialogSaveAs)
        .Title = "选择保存位置"
        .InitialFileName = IIf(fileName = "", ActivePresentation.FullName, fileName)
        .FilterIndex = IIf(LCase(format) = ".wmv", 17, 16) ' MP4
        '.Filters.Clear ' 都不支持
        '.Filters.Add "MP4 视频文件", "*.mp4", 0
        '.Filters.Add "WMV 视频文件", "*.wmv", 1
        '.Filters.Add "所有文件", "*", 2
        If .Show = -1 Then
            SaveAsFile = .SelectedItems(1)
        Else
            SaveAsFile = False
        End If
    End With
End Function

Private Function GetExtension(fileName As String) As String()
    Dim divide(1) As String, directories() As String, separators() As String
    fileName = Replace(fileName, "/", "\")
    directories = Split(fileName, "\")
    separators = Split(directories(UBound(directories)), ".")
    divide(1) = "." & IIf(UBound(separators) = 0, "", separators(UBound(separators)))
    If UBound(separators) <> 0 Then ReDim Preserve separators(LBound(separators) To UBound(separators) - 1)
    directories(UBound(directories)) = Join(separators, ".")
    divide(0) = Join(directories, "\")
    GetExtension = divide
End Function

Private Sub SetFileNames(fileName As String)
    Dim fileNames() As String
    fileNames = GetExtension(fileName)
    FileNameText.Text = fileNames(0)
    For i = LBound(FormatCombo.List) To UBound(FormatCombo.List)
        Dim extension As String
        extension = FormatCombo.List(i)
        If extension = fileNames(1) Then
            FormatCombo.Text = extension
            Exit For
        End If
    Next
End Sub

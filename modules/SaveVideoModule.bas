Attribute VB_Name = "SaveVideoModule"
Public Sub SaveVideo( _
    name As String, _
    format As String, _
    useNarrations As Boolean, _
    duration As Long, _
    height As Long, _
    fps As Long, _
    quality As Long _
)
    ActivePresentation.CreateVideo _
        name & format, _
        useNarrations, _
        duration, _
        height, _
        fps, _
        quality
End Sub

Public Sub OutputVideoAction()
    OutputVideoForm.Show
End Sub

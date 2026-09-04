Attribute VB_Name = "M_ViewsCamera"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)
    Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal virtualKey As Long) As Integer
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal virtualKey As Long) As Integer
#End If

Private Const VK_ESCAPE As Long = 27

Public Sub TileAndSynchronizeResultViews(ByVal removedWindow As Object, ByVal addedWindow As Object)
    Dim activeWindow As Object

    If CATIA.Windows.Count <> 2 Then
        Err.Raise vbObjectError + 1790, "TileAndSynchronizeResultViews", _
                  "Exactly two CATIA document windows are required for the locked split view."
    End If
    CATIA.Windows.Arrange catArrangeTiledVertical
    DoEvents

    On Error Resume Next
    removedWindow.ActiveViewer.Reframe
    removedWindow.ActiveViewer.Update
    CopyViewpoint removedWindow, addedWindow
    addedWindow.ActiveViewer.Update
    Err.Clear
    On Error GoTo 0

    MsgBox "The Removed and Added result products are tiled vertically." & vbCrLf & vbCrLf & _
           "Camera orientation and zoom will now stay synchronized while this macro is running." & _
           vbCrLf & "Press Esc when you want to release the camera lock.", _
           vbInformation, "Grid Compare - Camera Lock"

    ' Do not let Esc used to dismiss the information box immediately release
    ' the lock. Wait until that physical key has been released.
    Do While EscapeIsDownForView()
        DoEvents
        Sleep 25
    Loop
    SetStatus "Grid Compare camera lock active - press Esc to release"

    Do
        DoEvents
        If EscapeIsDownForView() Then Exit Do

        Set activeWindow = Nothing
        On Error Resume Next
        Set activeWindow = CATIA.ActiveWindow
        On Error GoTo CameraEnded

        If Not activeWindow Is Nothing Then
            If WindowMatches(activeWindow, removedWindow) Then
                CopyViewpoint removedWindow, addedWindow
            ElseIf WindowMatches(activeWindow, addedWindow) Then
                CopyViewpoint addedWindow, removedWindow
            End If
        End If
        Sleep 75
    Loop

CameraEnded:
    SetStatus "Grid Compare complete - camera lock released"
End Sub

Private Function WindowMatches(ByVal firstWindow As Object, ByVal secondWindow As Object) As Boolean
    On Error GoTo NotSame
    WindowMatches = (StrComp(firstWindow.Caption, secondWindow.Caption, vbBinaryCompare) = 0)
    Exit Function
NotSame:
    WindowMatches = False
    Err.Clear
End Function

Private Sub CopyViewpoint(ByVal sourceWindow As Object, ByVal targetWindow As Object)
    Dim sourceViewer As Object
    Dim targetViewer As Object
    Dim sourceViewpoint As Object
    Dim targetViewpoint As Object
    Dim coordinates() As Variant

    On Error GoTo CopyFailed
    Set sourceViewer = sourceWindow.ActiveViewer
    Set targetViewer = targetWindow.ActiveViewer
    Set sourceViewpoint = sourceViewer.Viewpoint3D
    Set targetViewpoint = targetViewer.Viewpoint3D

    ReDim coordinates(2)
    sourceViewpoint.GetOrigin coordinates
    targetViewpoint.PutOrigin coordinates
    sourceViewpoint.GetSightDirection coordinates
    targetViewpoint.PutSightDirection coordinates
    sourceViewpoint.GetUpDirection coordinates
    targetViewpoint.PutUpDirection coordinates

    targetViewpoint.ProjectionMode = sourceViewpoint.ProjectionMode
    targetViewpoint.FocusDistance = sourceViewpoint.FocusDistance
    On Error Resume Next
    targetViewpoint.Zoom = sourceViewpoint.Zoom
    targetViewpoint.FieldOfView = sourceViewpoint.FieldOfView
    Err.Clear
    On Error GoTo CopyFailed
    targetViewer.Update
    Exit Sub

CopyFailed:
    Err.Clear
End Sub

Private Function EscapeIsDownForView() As Boolean
    EscapeIsDownForView = ((GetAsyncKeyState(VK_ESCAPE) And &H8000) <> 0)
End Function

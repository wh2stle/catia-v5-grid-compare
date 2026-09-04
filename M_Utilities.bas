Attribute VB_Name = "M_Utilities"
Option Explicit

Public gRunId As String
Public gTempFolder As String
Public gLogText As String
Public gCancelRequested As Boolean

Public Sub InitializeRunState()
    gRunId = Format$(Now, "yyyymmdd_hhnnss")
    gLogText = "CATIA V5 Grid Compare" & vbCrLf
    gLogText = gLogText & "Run ID: " & gRunId & vbCrLf
    gLogText = gLogText & "Started: " & CStr(Now) & vbCrLf
    gCancelRequested = False
End Sub

Public Sub LogMessage(ByVal message As String)
    gLogText = gLogText & message & vbCrLf
End Sub

Public Sub SetStatus(ByVal message As String)
    On Error Resume Next
    CATIA.StatusBar = message
    On Error GoTo 0
End Sub

Public Function JoinPath(ByVal folderPath As String, ByVal leafName As String) As String
    If Right$(folderPath, 1) = "\" Then
        JoinPath = folderPath & leafName
    Else
        JoinPath = folderPath & "\" & leafName
    End If
End Function

Public Function SafeFeatureName(ByVal rawName As String) As String
    Dim result As String
    Dim badChars As Variant
    Dim i As Long

    result = Trim$(rawName)
    badChars = Array("\", "/", ":", "*", "?", Chr$(34), "<", ">", "|", ".", ";", ",")
    For i = LBound(badChars) To UBound(badChars)
        result = Replace$(result, CStr(badChars(i)), "_")
    Next i
    If Len(result) = 0 Then result = "Unnamed"
    If Len(result) > 60 Then result = Left$(result, 60)
    SafeFeatureName = result
End Function

Public Function FileLeaf(ByVal fullPath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileLeaf = fso.GetFileName(fullPath)
End Function

Public Function FileStem(ByVal fullPath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileStem = fso.GetBaseName(fullPath)
End Function

Public Function IsSupportedCATIAPath(ByVal fullPath As String) As Boolean
    Dim fso As Object
    Dim extensionName As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    extensionName = LCase$(fso.GetExtensionName(fullPath))
    IsSupportedCATIAPath = (extensionName = "catpart" Or extensionName = "catproduct")
End Function

Public Function MaxDouble(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxDouble = a Else MaxDouble = b
End Function

Public Function MinDouble(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinDouble = a Else MinDouble = b
End Function

Public Function BodyHasShape(ByVal bodyObject As Object) As Boolean
    On Error GoTo NotAvailable
    BodyHasShape = (bodyObject.Shapes.Count > 0)
    Exit Function
NotAvailable:
    BodyHasShape = False
End Function

Public Function MeasureBodyVolumeM3(ByVal partDocument As Object, ByVal bodyObject As Object) As Double
    Dim volumeErrorNumber As Long
    Dim volumeErrorDescription As String

    On Error GoTo NoVolume
    ' Product.Analyze is part of Product Structure automation and does not invoke
    ' the DMU Space Analysis workbench. Every caller supplies the main result body
    ' of a temporary CATPart, so the analyzed product volume is that body's result.
    partDocument.Part.InWorkObject = bodyObject
    partDocument.Part.UpdateObject bodyObject
    MeasureBodyVolumeM3 = CDbl(partDocument.Product.Analyze.Volume)
    Exit Function
NoVolume:
    volumeErrorNumber = Err.Number
    volumeErrorDescription = Err.Description
    Err.Clear
    Err.Raise vbObjectError + 1702, "MeasureBodyVolumeM3", _
              "CATIA could not analyze the main result body's volume (CATIA error " & _
              CStr(volumeErrorNumber) & "): " & volumeErrorDescription
End Function

Public Sub DeleteObjectFromDocument(ByVal targetDocument As Object, ByVal targetObject As Object)
    Dim selectionObject As Object

    If targetObject Is Nothing Then Exit Sub
    Set selectionObject = targetDocument.Selection
    selectionObject.Clear
    selectionObject.Add targetObject
    selectionObject.Delete
    selectionObject.Clear
End Sub

Public Sub SetAppearance(ByVal targetDocument As Object, ByVal targetObject As Object, _
                         ByVal redValue As Long, ByVal greenValue As Long, ByVal blueValue As Long, _
                         ByVal opacityValue As Long)
    Dim selectionObject As Object
    Dim visualProperties As Object
    Dim appearanceErrorNumber As Long
    Dim appearanceErrorDescription As String
    Dim targetName As String

    On Error GoTo AppearanceFailed
    Set selectionObject = targetDocument.Selection
    selectionObject.Clear
    selectionObject.Add targetObject
    Set visualProperties = selectionObject.VisProperties
    visualProperties.SetRealColor redValue, greenValue, blueValue, 1
    visualProperties.SetRealOpacity opacityValue, 1
    visualProperties.SetShow catVisPropertyShowAttr
    selectionObject.Clear
    Exit Sub
AppearanceFailed:
    appearanceErrorNumber = Err.Number
    appearanceErrorDescription = Err.Description
    Err.Clear
    targetName = "selected result object"
    On Error Resume Next
    targetName = targetObject.Name
    targetDocument.Selection.Clear
    Err.Clear
    On Error GoTo 0
    Err.Raise vbObjectError + 1703, "SetAppearance", _
              "Could not apply the required color/opacity to '" & targetName & _
              "' (CATIA error " & CStr(appearanceErrorNumber) & "): " & _
              appearanceErrorDescription
End Sub

Public Sub SetVisibility(ByVal targetDocument As Object, ByVal targetObject As Object, ByVal showValue As Long)
    Dim selectionObject As Object

    On Error Resume Next
    Set selectionObject = targetDocument.Selection
    selectionObject.Clear
    selectionObject.Add targetObject
    selectionObject.VisProperties.SetShow showValue
    selectionObject.Clear
    On Error GoTo 0
End Sub

Public Function CountMarkedCells(ByRef flags() As Boolean) As Long
    Dim i As Long
    Dim result As Long

    For i = LBound(flags) To UBound(flags)
        If flags(i) Then result = result + 1
    Next i
    CountMarkedCells = result
End Function

Public Sub WriteUtf8TextFile(ByVal fullPath As String, ByVal contents As String)
    Dim streamObject As Object

    Set streamObject = CreateObject("ADODB.Stream")
    streamObject.Type = 2
    streamObject.Charset = "utf-8"
    streamObject.Open
    streamObject.WriteText contents
    streamObject.SaveToFile fullPath, 2
    streamObject.Close
End Sub

Public Sub SafeCloseDocument(ByVal documentObject As Object)
    If documentObject Is Nothing Then Exit Sub
    On Error Resume Next
    documentObject.Close
    Err.Clear
    On Error GoTo 0
End Sub

Public Sub CloseRunOwnedDocuments()
    Dim documentIndex As Long
    Dim documentObject As Object
    Dim partNumber As String

    If Len(gRunId) = 0 Then Exit Sub
    For documentIndex = CATIA.Documents.Count To 1 Step -1
        Set documentObject = CATIA.Documents.Item(documentIndex)
        partNumber = ""
        On Error Resume Next
        Err.Clear
        partNumber = CStr(documentObject.Product.PartNumber)
        Err.Clear
        On Error GoTo 0

        If IsRunOwnedPartNumber(partNumber) Then SafeCloseDocument documentObject
    Next documentIndex
End Sub

Private Function IsRunOwnedPartNumber(ByVal partNumber As String) As Boolean
    Dim prefixes As Variant
    Dim prefixValue As Variant

    prefixes = Array( _
        "GRID_COMPARE_SOURCE_", "FLAT_REVISION_1_", "FLAT_REVISED_", _
        "ADDED_RAW_", "REMOVED_RAW_", "GRID_CONSTRUCTION_", _
        "ADDED_CELLS_", "REMOVED_CELLS_", "ADDED_VIEW_", "REMOVED_VIEW_")
    For Each prefixValue In prefixes
        If StrComp(partNumber, CStr(prefixValue) & gRunId, vbBinaryCompare) = 0 Then
            IsRunOwnedPartNumber = True
            Exit Function
        End If
    Next prefixValue
End Function

Public Sub SaveDocumentAs(ByVal documentObject As Object, ByVal fullPath As String)
    On Error GoTo SaveFailed
    documentObject.SaveAs fullPath
    LogMessage "Saved: " & fullPath
    Exit Sub
SaveFailed:
    Err.Raise vbObjectError + 1701, "SaveDocumentAs", _
              "CATIA could not save '" & fullPath & "'. " & Err.Description
End Sub

Public Function FormatInvariant(ByVal numericValue As Double, Optional ByVal decimalPlaces As Long = 6) As String
    Dim result As String
    result = Format$(numericValue, "0." & String$(decimalPlaces, "0"))
    result = Replace$(result, ",", ".")
    FormatInvariant = result
End Function

Public Sub PreparePrivateTempFolder()
    Dim fso As Object
    Dim tempRoot As String
    Dim rootFolder As Object
    Dim subFolder As Object
    Dim oldPaths As Collection
    Dim oldPath As Variant

    Set fso = CreateObject("Scripting.FileSystemObject")
    tempRoot = Environ$("TEMP")
    If Len(tempRoot) = 0 Then tempRoot = Environ$("TMP")
    If Len(tempRoot) = 0 Then Exit Sub

    ' Delete only folders created by this macro, never CATIA's shared cache.
    Set oldPaths = New Collection
    On Error Resume Next
    Set rootFolder = fso.GetFolder(tempRoot)
    For Each subFolder In rootFolder.SubFolders
        If Left$(subFolder.Name, Len(TEMP_PREFIX)) = TEMP_PREFIX Then
            oldPaths.Add subFolder.Path
        End If
    Next subFolder
    On Error GoTo 0

    For Each oldPath In oldPaths
        On Error Resume Next
        fso.DeleteFolder CStr(oldPath), True
        On Error GoTo 0
    Next oldPath

    gTempFolder = JoinPath(tempRoot, TEMP_PREFIX & gRunId)
    If Not fso.FolderExists(gTempFolder) Then fso.CreateFolder gTempFolder
End Sub

Public Sub RemovePrivateTempFolder()
    Dim fso As Object

    If Len(gTempFolder) = 0 Then Exit Sub
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    If fso.FolderExists(gTempFolder) Then fso.DeleteFolder gTempFolder, True
    On Error GoTo 0
End Sub

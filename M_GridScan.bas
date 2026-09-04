Attribute VB_Name = "M_GridScan"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal virtualKey As Long) As Integer
#Else
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal virtualKey As Long) As Integer
#End If

Private Const VK_ESCAPE As Long = 27

Public Sub ScanDifferenceCells(ByVal differenceDocument As Object, _
                               ByVal differenceBody As Object, _
                               ByRef gridSpec As TGridSpec, _
                               ByRef flags() As Boolean, _
                               ByRef probeCount As Long, _
                               ByVal phaseLabel As String)
    Dim totalCells As Long
    Dim baselineVolumeM3 As Double
    Dim finalVolumeM3 As Double

    differenceDocument.Activate
    totalCells = gridSpec.NX * gridSpec.NY * gridSpec.NZ
    ReDim flags(0 To totalCells - 1)
    probeCount = 0
    baselineVolumeM3 = MeasureBodyVolumeM3(differenceDocument, differenceBody)

    If baselineVolumeM3 < MIN_CHANGE_VOLUME_M3 Then
        LogMessage phaseLabel & ": raw difference is below the configured noise floor."
        Exit Sub
    End If

    gCancelRequested = False
    ScanIndexRange differenceDocument, differenceBody, gridSpec, flags, probeCount, _
                   phaseLabel, baselineVolumeM3, _
                   0, gridSpec.NX, 0, gridSpec.NY, 0, gridSpec.NZ

    If gCancelRequested Then
        Err.Raise vbObjectError + 1760, "ScanDifferenceCells", _
                  phaseLabel & " scan canceled with Esc."
    End If
    finalVolumeM3 = MeasureBodyVolumeM3(differenceDocument, differenceBody)
    If Abs(finalVolumeM3 - baselineVolumeM3) > MIN_CHANGE_VOLUME_M3 Then
        Err.Raise vbObjectError + 1764, "ScanDifferenceCells", _
                  phaseLabel & " difference body was not restored after the final grid probe."
    End If
    LogMessage phaseLabel & ": " & CStr(CountMarkedCells(flags)) & _
               " marked cell(s), " & CStr(probeCount) & " Boolean probe(s)."
End Sub

Private Sub ScanIndexRange(ByVal differenceDocument As Object, _
                           ByVal differenceBody As Object, _
                           ByRef gridSpec As TGridSpec, _
                           ByRef flags() As Boolean, _
                           ByRef probeCount As Long, _
                           ByVal phaseLabel As String, _
                           ByVal baselineVolumeM3 As Double, _
                           ByVal iStart As Long, ByVal iEnd As Long, _
                           ByVal jStart As Long, ByVal jEnd As Long, _
                           ByVal kStart As Long, ByVal kEnd As Long)
    Dim rangeBounds As TBounds
    Dim intersectionVolumeM3 As Double
    Dim blockVolumeM3 As Double
    Dim countI As Long
    Dim countJ As Long
    Dim countK As Long
    Dim middleIndex As Long
    Dim spanI As Double
    Dim spanJ As Double
    Dim spanK As Double

    If gCancelRequested Then Exit Sub
    probeCount = probeCount + 1
    If (probeCount Mod PROGRESS_INTERVAL) = 0 Then
        DoEvents
        If EscapeIsDown() Then
            gCancelRequested = True
            Exit Sub
        End If
        SetStatus phaseLabel & " cell scan: " & Format$(probeCount, "#,##0") & _
                  " Boolean probes; Esc cancels"
    End If

    BoundsForIndexRange gridSpec, iStart, iEnd, jStart, jEnd, kStart, kEnd, rangeBounds
    intersectionVolumeM3 = ProbeIntersectionVolume(differenceDocument, differenceBody, _
                                                   rangeBounds, baselineVolumeM3, probeCount)
    ' Once the whole difference passes the run-level noise floor, keep the
    ' requested any-portion rule for individual cells. CATIA's Boolean kernel
    ' decides whether the intersection is geometrically valid; any positive
    ' analyzed volume is retained.
    If intersectionVolumeM3 <= EMPTY_VOLUME_EPS_M3 Then Exit Sub

    blockVolumeM3 = ((rangeBounds.MaxX - rangeBounds.MinX) * _
                     (rangeBounds.MaxY - rangeBounds.MinY) * _
                     (rangeBounds.MaxZ - rangeBounds.MinZ)) / 1000000000#
    ' Never use a "nearly full" relative shortcut: a small real cavity could
    ' otherwise cause completely empty leaf cells to be painted. If rounding
    ' puts the value just below full volume, subdivision is slower but correct.
    If intersectionVolumeM3 >= blockVolumeM3 Then
        MarkIndexRange gridSpec, flags, iStart, iEnd, jStart, jEnd, kStart, kEnd
        Exit Sub
    End If

    countI = iEnd - iStart
    countJ = jEnd - jStart
    countK = kEnd - kStart
    If countI = 1 And countJ = 1 And countK = 1 Then
        flags(CellLinearIndex(gridSpec, iStart, jStart, kStart)) = True
        Exit Sub
    End If

    ' Exclude axes that are already one cell wide. Without this guard, unequal
    ' physical cell sizes could select an unsplittable axis and recurse forever.
    spanI = -1#
    spanJ = -1#
    spanK = -1#
    If countI > 1 Then spanI = CDbl(countI) * gridSpec.DX
    If countJ > 1 Then spanJ = CDbl(countJ) * gridSpec.DY
    If countK > 1 Then spanK = CDbl(countK) * gridSpec.DZ

    If spanI >= spanJ And spanI >= spanK Then
        middleIndex = iStart + countI \ 2
        ScanIndexRange differenceDocument, differenceBody, gridSpec, flags, probeCount, _
                       phaseLabel, baselineVolumeM3, iStart, middleIndex, jStart, jEnd, kStart, kEnd
        ScanIndexRange differenceDocument, differenceBody, gridSpec, flags, probeCount, _
                       phaseLabel, baselineVolumeM3, middleIndex, iEnd, jStart, jEnd, kStart, kEnd
    ElseIf spanJ >= spanK Then
        middleIndex = jStart + countJ \ 2
        ScanIndexRange differenceDocument, differenceBody, gridSpec, flags, probeCount, _
                       phaseLabel, baselineVolumeM3, iStart, iEnd, jStart, middleIndex, kStart, kEnd
        ScanIndexRange differenceDocument, differenceBody, gridSpec, flags, probeCount, _
                       phaseLabel, baselineVolumeM3, iStart, iEnd, middleIndex, jEnd, kStart, kEnd
    Else
        middleIndex = kStart + countK \ 2
        ScanIndexRange differenceDocument, differenceBody, gridSpec, flags, probeCount, _
                       phaseLabel, baselineVolumeM3, iStart, iEnd, jStart, jEnd, kStart, middleIndex
        ScanIndexRange differenceDocument, differenceBody, gridSpec, flags, probeCount, _
                       phaseLabel, baselineVolumeM3, iStart, iEnd, jStart, jEnd, middleIndex, kEnd
    End If
End Sub

Private Function ProbeIntersectionVolume(ByVal differenceDocument As Object, _
                                         ByVal differenceBody As Object, _
                                         ByRef probeBounds As TBounds, _
                                         ByVal baselineVolumeM3 As Double, _
                                         ByVal probeNumber As Long) As Double
    Dim partObject As Object
    Dim supportSet As Object
    Dim planeObject As Object
    Dim probeBody As Object
    Dim intersectionOperation As Object
    Dim shapeFactory As Object
    Dim booleanWorked As Boolean
    Dim restoredVolumeM3 As Double
    Dim savedErrorNumber As Long
    Dim savedErrorDescription As String

    Set partObject = differenceDocument.Part
    On Error GoTo GeometryFailed
    Set supportSet = partObject.HybridBodies.Add
    supportSet.Name = "__PROBE_SUPPORT__"
    Set planeObject = CreateOffsetXYPlane(differenceDocument, supportSet, probeBounds.MinZ, "__PROBE_PLANE__")
    Set probeBody = CreateRectangularPrismBody(differenceDocument, planeObject, _
                    probeBounds.MinX, probeBounds.MinY, probeBounds.MaxX, probeBounds.MaxY, _
                    probeBounds.MaxZ - probeBounds.MinZ, "__PROBE_BODY__")

    Set shapeFactory = partObject.ShapeFactory
    partObject.InWorkObject = differenceBody
    On Error Resume Next
    Err.Clear
    Set intersectionOperation = shapeFactory.AddNewIntersect(probeBody)
    partObject.UpdateObject intersectionOperation
    booleanWorked = (Err.Number = 0)
    Err.Clear
    On Error GoTo GeometryFailed

    If booleanWorked Then
        ProbeIntersectionVolume = MeasureBodyVolumeM3(differenceDocument, differenceBody)
    Else
        ' Disjoint Part Design intersections normally report an update error.
        ProbeIntersectionVolume = 0#
    End If

    CleanupProbe differenceDocument, intersectionOperation, probeBody, supportSet

    If (probeNumber Mod PROGRESS_INTERVAL) = 0 Then
        restoredVolumeM3 = MeasureBodyVolumeM3(differenceDocument, differenceBody)
        If Abs(restoredVolumeM3 - baselineVolumeM3) > MIN_CHANGE_VOLUME_M3 Then
            Err.Raise vbObjectError + 1761, "ProbeIntersectionVolume", _
                      "A temporary Boolean probe did not restore the difference body safely."
        End If
    End If
    Exit Function

GeometryFailed:
    savedErrorNumber = Err.Number
    savedErrorDescription = Err.Description
    Err.Clear
    On Error Resume Next
    CleanupProbe differenceDocument, intersectionOperation, probeBody, supportSet
    On Error GoTo 0
    Err.Raise vbObjectError + 1762, "ProbeIntersectionVolume", _
              "Could not build a temporary grid probe. CATIA error " & _
              CStr(savedErrorNumber) & ": " & savedErrorDescription
End Function

Private Sub CleanupProbe(ByVal partDocument As Object, _
                         ByVal intersectionOperation As Object, _
                         ByVal probeBody As Object, _
                         ByVal supportSet As Object)
    Dim cleanupError As Long
    Dim cleanupDescription As String
    Dim objectStillExists As Boolean
    Dim objectName As String

    On Error GoTo CleanupFailed
    If Not intersectionOperation Is Nothing Then
        DeleteObjectFromDocument partDocument, intersectionOperation
        partDocument.Part.Update
    End If

    ' Depending on the V5 service pack, deleting a Part Design Boolean may also
    ' delete its aggregated tool Body. Only issue a second delete if that COM
    ' object is still accessible; an accessible object must delete successfully.
    If Not probeBody Is Nothing Then
        objectStillExists = False
        On Error Resume Next
        Err.Clear
        objectName = probeBody.Name
        objectStillExists = (Err.Number = 0)
        Err.Clear
        On Error GoTo CleanupFailed
        If objectStillExists Then DeleteObjectFromDocument partDocument, probeBody
    End If

    If Not supportSet Is Nothing Then
        objectStillExists = False
        On Error Resume Next
        Err.Clear
        objectName = supportSet.Name
        objectStillExists = (Err.Number = 0)
        Err.Clear
        On Error GoTo CleanupFailed
        If objectStillExists Then DeleteObjectFromDocument partDocument, supportSet
    End If
    partDocument.Part.Update
    Exit Sub

CleanupFailed:
    cleanupError = Err.Number
    cleanupDescription = Err.Description
    Err.Clear
    Err.Raise vbObjectError + 1763, "CleanupProbe", _
              "Could not remove temporary Boolean geometry (" & CStr(cleanupError) & "): " & _
              cleanupDescription
End Sub

Private Sub BoundsForIndexRange(ByRef gridSpec As TGridSpec, _
                                ByVal iStart As Long, ByVal iEnd As Long, _
                                ByVal jStart As Long, ByVal jEnd As Long, _
                                ByVal kStart As Long, ByVal kEnd As Long, _
                                ByRef resultBounds As TBounds)
    resultBounds.MinX = gridSpec.Bounds.MinX + CDbl(iStart) * gridSpec.DX
    resultBounds.MaxX = gridSpec.Bounds.MinX + CDbl(iEnd) * gridSpec.DX
    resultBounds.MinY = gridSpec.Bounds.MinY + CDbl(jStart) * gridSpec.DY
    resultBounds.MaxY = gridSpec.Bounds.MinY + CDbl(jEnd) * gridSpec.DY
    resultBounds.MinZ = gridSpec.Bounds.MinZ + CDbl(kStart) * gridSpec.DZ
    resultBounds.MaxZ = gridSpec.Bounds.MinZ + CDbl(kEnd) * gridSpec.DZ
End Sub

Private Sub MarkIndexRange(ByRef gridSpec As TGridSpec, _
                           ByRef flags() As Boolean, _
                           ByVal iStart As Long, ByVal iEnd As Long, _
                           ByVal jStart As Long, ByVal jEnd As Long, _
                           ByVal kStart As Long, ByVal kEnd As Long)
    Dim i As Long
    Dim j As Long
    Dim k As Long

    For k = kStart To kEnd - 1
        For j = jStart To jEnd - 1
            For i = iStart To iEnd - 1
                flags(CellLinearIndex(gridSpec, i, j, k)) = True
            Next i
        Next j
    Next k
End Sub

Public Function CellLinearIndex(ByRef gridSpec As TGridSpec, _
                                ByVal i As Long, ByVal j As Long, ByVal k As Long) As Long
    CellLinearIndex = i + gridSpec.NX * (j + gridSpec.NY * k)
End Function

Private Function EscapeIsDown() As Boolean
    EscapeIsDown = ((GetAsyncKeyState(VK_ESCAPE) And &H8000) <> 0)
End Function

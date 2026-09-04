Attribute VB_Name = "M_Bounds"
Option Explicit

Public Function ComputeUnionGrid(ByVal originalFlatDocument As Object, _
                                 ByVal revisedFlatDocument As Object, _
                                 ByRef unionWithoutMargin As TBounds, _
                                 ByRef gridSpec As TGridSpec) As Boolean
    Dim originalBounds As TBounds
    Dim revisedBounds As TBounds
    Dim spanX As Double
    Dim spanY As Double
    Dim spanZ As Double

    SetStatus "Computing the union bounding box..."
    If Not ComputePartBounds(originalFlatDocument, originalBounds) Then Exit Function
    If Not ComputePartBounds(revisedFlatDocument, revisedBounds) Then Exit Function

    unionWithoutMargin.MinX = MinDouble(originalBounds.MinX, revisedBounds.MinX)
    unionWithoutMargin.MinY = MinDouble(originalBounds.MinY, revisedBounds.MinY)
    unionWithoutMargin.MinZ = MinDouble(originalBounds.MinZ, revisedBounds.MinZ)
    unionWithoutMargin.MaxX = MaxDouble(originalBounds.MaxX, revisedBounds.MaxX)
    unionWithoutMargin.MaxY = MaxDouble(originalBounds.MaxY, revisedBounds.MaxY)
    unionWithoutMargin.MaxZ = MaxDouble(originalBounds.MaxZ, revisedBounds.MaxZ)

    spanX = unionWithoutMargin.MaxX - unionWithoutMargin.MinX
    spanY = unionWithoutMargin.MaxY - unionWithoutMargin.MinY
    spanZ = unionWithoutMargin.MaxZ - unionWithoutMargin.MinZ

    If spanX <= 0# Or spanY <= 0# Or spanZ <= 0# Then
        Err.Raise vbObjectError + 1730, "ComputeUnionGrid", _
                  "The union is degenerate on at least one axis. Solid 3D input is required."
    End If

    gridSpec.Bounds.MinX = unionWithoutMargin.MinX - spanX * BOUNDS_MARGIN_FRACTION
    gridSpec.Bounds.MinY = unionWithoutMargin.MinY - spanY * BOUNDS_MARGIN_FRACTION
    gridSpec.Bounds.MinZ = unionWithoutMargin.MinZ - spanZ * BOUNDS_MARGIN_FRACTION
    gridSpec.Bounds.MaxX = unionWithoutMargin.MaxX + spanX * BOUNDS_MARGIN_FRACTION
    gridSpec.Bounds.MaxY = unionWithoutMargin.MaxY + spanY * BOUNDS_MARGIN_FRACTION
    gridSpec.Bounds.MaxZ = unionWithoutMargin.MaxZ + spanZ * BOUNDS_MARGIN_FRACTION

    gridSpec.NX = GRID_DIVISIONS_X
    gridSpec.NY = GRID_DIVISIONS_Y
    gridSpec.NZ = GRID_DIVISIONS_Z
    gridSpec.DX = (gridSpec.Bounds.MaxX - gridSpec.Bounds.MinX) / CDbl(gridSpec.NX)
    gridSpec.DY = (gridSpec.Bounds.MaxY - gridSpec.Bounds.MinY) / CDbl(gridSpec.NY)
    gridSpec.DZ = (gridSpec.Bounds.MaxZ - gridSpec.Bounds.MinZ) / CDbl(gridSpec.NZ)

    LogBounds "Original bounds", originalBounds
    LogBounds "Revised bounds", revisedBounds
    LogBounds "Union before margin", unionWithoutMargin
    LogBounds "Grid bounds (" & _
              FormatInvariant(BOUNDS_MARGIN_FRACTION * 100#, 3) & "% each side)", _
              gridSpec.Bounds
    LogMessage "Cell size mm: " & FormatInvariant(gridSpec.DX, 6) & " x " & _
               FormatInvariant(gridSpec.DY, 6) & " x " & FormatInvariant(gridSpec.DZ, 6)

    ComputeUnionGrid = True
End Function

Public Function ComputePartBounds(ByVal partDocument As Object, ByRef resultBounds As TBounds) As Boolean
    Dim bodyIndex As Long
    Dim bodyObject As Object
    Dim temporarySet As Object
    Dim coordinateValue As Double
    Dim foundAny As Boolean
    Dim failedBodyName As String
    Dim extremumError As String

    partDocument.Activate
    resultBounds.MinX = 1E+99
    resultBounds.MinY = 1E+99
    resultBounds.MinZ = 1E+99
    resultBounds.MaxX = -1E+99
    resultBounds.MaxY = -1E+99
    resultBounds.MaxZ = -1E+99

    Set temporarySet = partDocument.Part.HybridBodies.Add
    temporarySet.Name = "__GRID_COMPARE_BOUNDING_PROBES__"

    For bodyIndex = 1 To partDocument.Part.Bodies.Count
        Set bodyObject = partDocument.Part.Bodies.Item(bodyIndex)
        If BodyHasShape(bodyObject) Then
            failedBodyName = bodyObject.Name
            extremumError = ""
            If Not TryBodyExtremum(partDocument, bodyObject, temporarySet, _
                                   1#, 0#, 0#, 0, coordinateValue, extremumError) Then GoTo BoundsFailed
            resultBounds.MinX = MinDouble(resultBounds.MinX, coordinateValue)

            If Not TryBodyExtremum(partDocument, bodyObject, temporarySet, _
                                   1#, 0#, 0#, 1, coordinateValue, extremumError) Then GoTo BoundsFailed
            resultBounds.MaxX = MaxDouble(resultBounds.MaxX, coordinateValue)

            If Not TryBodyExtremum(partDocument, bodyObject, temporarySet, _
                                   0#, 1#, 0#, 0, coordinateValue, extremumError) Then GoTo BoundsFailed
            resultBounds.MinY = MinDouble(resultBounds.MinY, coordinateValue)

            If Not TryBodyExtremum(partDocument, bodyObject, temporarySet, _
                                   0#, 1#, 0#, 1, coordinateValue, extremumError) Then GoTo BoundsFailed
            resultBounds.MaxY = MaxDouble(resultBounds.MaxY, coordinateValue)

            If Not TryBodyExtremum(partDocument, bodyObject, temporarySet, _
                                   0#, 0#, 1#, 0, coordinateValue, extremumError) Then GoTo BoundsFailed
            resultBounds.MinZ = MinDouble(resultBounds.MinZ, coordinateValue)

            If Not TryBodyExtremum(partDocument, bodyObject, temporarySet, _
                                   0#, 0#, 1#, 1, coordinateValue, extremumError) Then GoTo BoundsFailed
            resultBounds.MaxZ = MaxDouble(resultBounds.MaxZ, coordinateValue)
            foundAny = True
        End If
    Next bodyIndex

    On Error Resume Next
    DeleteObjectFromDocument partDocument, temporarySet
    partDocument.Part.Update
    On Error GoTo 0

    If foundAny Then
        If resultBounds.MinX < 1E+98 And resultBounds.MinY < 1E+98 And resultBounds.MinZ < 1E+98 And _
           resultBounds.MaxX > -1E+98 And resultBounds.MaxY > -1E+98 And resultBounds.MaxZ > -1E+98 Then
            ComputePartBounds = True
        End If
    End If
    Exit Function

BoundsFailed:
    On Error Resume Next
    DeleteObjectFromDocument partDocument, temporarySet
    partDocument.Part.Update
    On Error GoTo 0
    Err.Raise vbObjectError + 1731, "ComputePartBounds", _
              "Could not create a complete bounding extremum for body '" & _
              failedBodyName & "'. " & extremumError
End Function

Private Function TryBodyExtremum(ByVal partDocument As Object, _
                                 ByVal bodyObject As Object, _
                                 ByVal temporarySet As Object, _
                                 ByVal directionX As Double, _
                                 ByVal directionY As Double, _
                                 ByVal directionZ As Double, _
                                 ByVal extremumType As Long, _
                                 ByRef coordinateValue As Double, _
                                 ByRef failureText As String) As Boolean
    Dim factoryObject As Object
    Dim directionObject As Object
    Dim secondaryDirection As Object
    Dim tertiaryDirection As Object
    Dim extremumObject As Object
    Dim bodyReference As Object
    Dim extremumReference As Object
    Dim coordinatePoint As Object
    Dim pointCoordinates() As Variant
    Dim extremumErrorNumber As Long
    Dim extremumErrorDescription As String

    On Error GoTo ExtremumFailed
    failureText = ""
    Set factoryObject = partDocument.Part.HybridShapeFactory
    Set bodyReference = partDocument.Part.CreateReferenceFromObject(bodyObject)
    Set directionObject = factoryObject.AddNewDirectionByCoord(directionX, directionY, directionZ)

    ' A one-direction extremum can be a whole planar face. Add two orthogonal
    ' tie-break directions so the result is point-like while preserving the
    ' requested primary minimum/maximum coordinate.
    If directionX <> 0# Then
        Set secondaryDirection = factoryObject.AddNewDirectionByCoord(0#, 1#, 0#)
        Set tertiaryDirection = factoryObject.AddNewDirectionByCoord(0#, 0#, 1#)
    ElseIf directionY <> 0# Then
        Set secondaryDirection = factoryObject.AddNewDirectionByCoord(1#, 0#, 0#)
        Set tertiaryDirection = factoryObject.AddNewDirectionByCoord(0#, 0#, 1#)
    Else
        Set secondaryDirection = factoryObject.AddNewDirectionByCoord(1#, 0#, 0#)
        Set tertiaryDirection = factoryObject.AddNewDirectionByCoord(0#, 1#, 0#)
    End If

    Set extremumObject = factoryObject.AddNewExtremum(bodyReference, directionObject, extremumType)
    extremumObject.Direction2 = secondaryDirection
    extremumObject.ExtremumType2 = 0
    extremumObject.Direction3 = tertiaryDirection
    extremumObject.ExtremumType3 = 0
    temporarySet.AppendHybridShape extremumObject
    partDocument.Part.UpdateObject extremumObject

    Set extremumReference = partDocument.Part.CreateReferenceFromObject(extremumObject)
    Set coordinatePoint = factoryObject.AddNewPointCoordWithReference(0#, 0#, 0#, extremumReference)
    temporarySet.AppendHybridShape coordinatePoint
    partDocument.Part.UpdateObject coordinatePoint
    ReDim pointCoordinates(2)
    coordinatePoint.GetCoordinates pointCoordinates

    If directionX <> 0# Then
        coordinateValue = pointCoordinates(0)
    ElseIf directionY <> 0# Then
        coordinateValue = pointCoordinates(1)
    Else
        coordinateValue = pointCoordinates(2)
    End If
    TryBodyExtremum = True
    Exit Function

ExtremumFailed:
    extremumErrorNumber = Err.Number
    extremumErrorDescription = Err.Description
    TryBodyExtremum = False
    Err.Clear
    failureText = "CATIA error " & CStr(extremumErrorNumber) & ": " & extremumErrorDescription
End Function

Public Sub LogBounds(ByVal labelText As String, ByRef boundsValue As TBounds)
    LogMessage labelText & ": X[" & FormatInvariant(boundsValue.MinX, 6) & ", " & _
               FormatInvariant(boundsValue.MaxX, 6) & "] Y[" & _
               FormatInvariant(boundsValue.MinY, 6) & ", " & _
               FormatInvariant(boundsValue.MaxY, 6) & "] Z[" & _
               FormatInvariant(boundsValue.MinZ, 6) & ", " & _
               FormatInvariant(boundsValue.MaxZ, 6) & "]"
End Sub

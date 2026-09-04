Attribute VB_Name = "M_BooleanDifference"
Option Explicit

' Builds base minus tool with native Part Design Booleans.  The result remains
' an unsaved working CATPart and is deleted after voxelization.
Public Function CreateDifferenceDocument(ByVal baseFlatDocument As Object, _
                                         ByVal toolFlatDocument As Object, _
                                         ByVal resultPartNumber As String, _
                                         ByRef differenceBody As Object, _
                                         ByRef resultVolumeM3 As Double, _
                                         ByRef resultIsEmpty As Boolean) As Object
    Dim resultDocument As Object
    Dim resultPart As Object
    Dim baseBody As Object
    Dim toolBody As Object
    Dim shapeFactory As Object
    Dim removeOperation As Object
    Dim baseVolumeBefore As Double
    Dim measureWorked As Boolean
    Dim measureErrorNumber As Long
    Dim measureErrorDescription As String

    SetStatus "Creating Boolean difference " & resultPartNumber & "..."
    Set resultDocument = CATIA.Documents.Add("Part")
    Set resultPart = resultDocument.Part
    resultDocument.Product.PartNumber = resultPartNumber

    Set baseBody = BuildUnifiedBody(baseFlatDocument, resultDocument, _
                                    resultPart.Bodies.Item(1), "BASE_UNION")
    baseVolumeBefore = MeasureBodyVolumeM3(resultDocument, baseBody)
    If baseVolumeBefore <= EMPTY_VOLUME_EPS_M3 Then
        Err.Raise vbObjectError + 1744, "CreateDifferenceDocument", _
                  "The base input for " & resultPartNumber & _
                  " has no positive solid volume. Surface-only input is not supported."
    End If

    Set toolBody = resultPart.Bodies.Add
    Set toolBody = BuildUnifiedBody(toolFlatDocument, resultDocument, toolBody, "TOOL_UNION")

    Set shapeFactory = resultPart.ShapeFactory
    resultPart.InWorkObject = baseBody

    On Error Resume Next
    Err.Clear
    Set removeOperation = shapeFactory.AddNewRemove(toolBody)
    resultPart.UpdateObject removeOperation
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        If DifferenceIsEmptyByIntersection(resultDocument, baseBody, toolBody, _
                                           removeOperation, baseVolumeBefore) Then
            resultIsEmpty = True
            resultVolumeM3 = 0#
            Set differenceBody = baseBody
            Set CreateDifferenceDocument = resultDocument
            LogMessage resultPartNumber & ": zero material after subtraction."
            Exit Function
        End If
        Err.Raise vbObjectError + 1740, "CreateDifferenceDocument", _
                  "CATIA could not update the Boolean Remove for " & resultPartNumber & "."
    End If
    On Error GoTo 0

    Set differenceBody = baseBody
    On Error Resume Next
    Err.Clear
    resultVolumeM3 = MeasureBodyVolumeM3(resultDocument, differenceBody)
    measureWorked = (Err.Number = 0)
    measureErrorNumber = Err.Number
    measureErrorDescription = Err.Description
    Err.Clear
    On Error GoTo 0

    If Not measureWorked Then
        ' Some V5 levels cannot analyze a successfully updated but empty
        ' PartBody. Restore the base and verify full tool containment instead
        ' of silently treating every analysis error as an empty result.
        If DifferenceIsEmptyByIntersection(resultDocument, baseBody, toolBody, _
                                           removeOperation, baseVolumeBefore) Then
            resultIsEmpty = True
            resultVolumeM3 = 0#
            Set differenceBody = baseBody
            Set CreateDifferenceDocument = resultDocument
            LogMessage resultPartNumber & ": zero material after subtraction."
            Exit Function
        End If
        Err.Raise vbObjectError + 1745, "CreateDifferenceDocument", _
                  "CATIA could not analyze the Boolean result (CATIA error " & _
                  CStr(measureErrorNumber) & "): " & measureErrorDescription
    End If

    resultIsEmpty = (resultVolumeM3 < MIN_CHANGE_VOLUME_M3)
    LogMessage resultPartNumber & " raw volume m^3: " & FormatInvariant(resultVolumeM3, 12)
    Set CreateDifferenceDocument = resultDocument
End Function

Private Function BuildUnifiedBody(ByVal sourceFlatDocument As Object, _
                                  ByVal targetDocument As Object, _
                                  ByVal accumulatorBody As Object, _
                                  ByVal bodyLabel As String) As Object
    Dim sourcePart As Object
    Dim sourceBody As Object
    Dim pastedBody As Object
    Dim sourceIndex As Long
    Dim pastedCount As Long

    Set sourcePart = sourceFlatDocument.Part
    accumulatorBody.Name = bodyLabel

    For sourceIndex = 1 To sourcePart.Bodies.Count
        Set sourceBody = sourcePart.Bodies.Item(sourceIndex)
        If BodyHasShape(sourceBody) Then
            If pastedCount = 0 Then
                Set pastedBody = accumulatorBody
            Else
                Set pastedBody = targetDocument.Part.Bodies.Add
                pastedBody.Name = bodyLabel & "_MEMBER_" & Format$(sourceIndex, "000")
            End If

            If Not PasteObjectAsResult(sourceFlatDocument, sourceBody, targetDocument, pastedBody) Then
                Err.Raise vbObjectError + 1741, "BuildUnifiedBody", _
                          "Could not copy a flattened body into " & bodyLabel & "."
            End If

            If pastedCount > 0 Then
                MergePositiveBody targetDocument, accumulatorBody, pastedBody
            End If
            pastedCount = pastedCount + 1
        End If
    Next sourceIndex

    If pastedCount = 0 Then
        Err.Raise vbObjectError + 1742, "BuildUnifiedBody", _
                  "No solid body was available for " & bodyLabel & "."
    End If
    targetDocument.Part.Update
    Set BuildUnifiedBody = accumulatorBody
End Function

Private Sub MergePositiveBody(ByVal partDocument As Object, _
                              ByVal accumulatorBody As Object, _
                              ByVal memberBody As Object)
    Dim shapeFactory As Object
    Dim booleanOperation As Object
    Dim operationWorked As Boolean

    Set shapeFactory = partDocument.Part.ShapeFactory
    partDocument.Part.InWorkObject = accumulatorBody

    On Error Resume Next
    Err.Clear
    Set booleanOperation = shapeFactory.AddNewAdd(memberBody)
    partDocument.Part.UpdateObject booleanOperation
    operationWorked = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0

    If operationWorked Then Exit Sub

    ' Add can reject some multi-domain cases. Assemble is the V5 fallback.
    On Error Resume Next
    If Not booleanOperation Is Nothing Then DeleteObjectFromDocument partDocument, booleanOperation
    partDocument.Part.Update
    Set booleanOperation = Nothing
    Err.Clear
    Set booleanOperation = shapeFactory.AddNewAssemble(memberBody)
    partDocument.Part.UpdateObject booleanOperation
    operationWorked = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0

    If Not operationWorked Then
        Err.Raise vbObjectError + 1743, "MergePositiveBody", _
                  "CATIA could not combine two flattened solid bodies."
    End If
End Sub

Private Function DifferenceIsEmptyByIntersection(ByVal partDocument As Object, _
                                                 ByVal baseBody As Object, _
                                                 ByVal toolBody As Object, _
                                                 ByVal failedRemove As Object, _
                                                 ByVal baseVolumeM3 As Double) As Boolean
    Dim intersectionOperation As Object
    Dim shapeFactory As Object
    Dim intersectionVolumeM3 As Double
    Dim toleranceM3 As Double
    Dim operationWorked As Boolean

    On Error Resume Next
    If Not failedRemove Is Nothing Then DeleteObjectFromDocument partDocument, failedRemove
    partDocument.Part.Update
    On Error GoTo 0

    Set shapeFactory = partDocument.Part.ShapeFactory
    partDocument.Part.InWorkObject = baseBody
    On Error Resume Next
    Err.Clear
    Set intersectionOperation = shapeFactory.AddNewIntersect(toolBody)
    partDocument.Part.UpdateObject intersectionOperation
    operationWorked = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
    If Not operationWorked Then Exit Function

    intersectionVolumeM3 = MeasureBodyVolumeM3(partDocument, baseBody)
    toleranceM3 = MIN_CHANGE_VOLUME_M3
    DifferenceIsEmptyByIntersection = (Abs(baseVolumeM3 - intersectionVolumeM3) <= toleranceM3)
End Function

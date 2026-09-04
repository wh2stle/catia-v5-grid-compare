Attribute VB_Name = "M_CopyFlatten"
Option Explicit

Private mLastFlattenFailure As String

' Creates a temporary CATPart containing disconnected, link-free result bodies.
' Product instances are copied from the assembly selection, so their instance
' transforms are baked into the result and the common assembly axis is retained.
Public Function FlattenSourceToPart(ByVal sourceDocument As Object, _
                                    ByVal resultPartNumber As String, _
                                    ByRef copiedCount As Long, _
                                    ByRef failedCount As Long) As Object
    Dim targetDocument As Object
    Dim targetPart As Object
    Dim firstBody As Object
    Dim sourceType As String
    Dim failureMessage As String

    copiedCount = 0
    failedCount = 0
    mLastFlattenFailure = ""
    sourceType = TypeName(sourceDocument)

    Set targetDocument = CATIA.Documents.Add("Part")
    Set targetPart = targetDocument.Part
    targetDocument.Product.PartNumber = resultPartNumber
    Set firstBody = targetPart.Bodies.Item(1)
    firstBody.Name = "FLAT_RESULT_001"

    If StrComp(sourceType, "PartDocument", vbTextCompare) = 0 Then
        FlattenStandalonePart sourceDocument, targetDocument, copiedCount, failedCount, firstBody
    ElseIf StrComp(sourceType, "ProductDocument", vbTextCompare) = 0 Then
        EnsureProductDesignMode sourceDocument.Product, sourceDocument.Name
        FlattenProductLevel sourceDocument.Product, sourceDocument, targetDocument, _
                            copiedCount, failedCount, firstBody
    Else
        failedCount = 1
        mLastFlattenFailure = "Unsupported CATIA document type: " & sourceType
        LogMessage "ERROR: " & mLastFlattenFailure
    End If

    If copiedCount = 0 Then
        failureMessage = "No solid component could be copied from " & sourceDocument.Name & "."
        If Len(mLastFlattenFailure) > 0 Then
            failureMessage = failureMessage & vbCrLf & mLastFlattenFailure
        End If
        Err.Raise vbObjectError + 1720, "FlattenSourceToPart", _
                  failureMessage
    End If

    targetPart.Update
    LogMessage resultPartNumber & ": copied " & CStr(copiedCount) & _
               " solid item(s); failed " & CStr(failedCount) & "."
    Set FlattenSourceToPart = targetDocument
End Function

Private Sub FlattenStandalonePart(ByVal sourcePartDocument As Object, _
                                  ByVal targetPartDocument As Object, _
                                  ByRef copiedCount As Long, _
                                  ByRef failedCount As Long, _
                                  ByVal firstBody As Object)
    Dim sourcePart As Object
    Dim sourceBody As Object
    Dim targetBody As Object
    Dim bodyIndex As Long
    Dim pasteFailure As Variant

    Set sourcePart = sourcePartDocument.Part
    For bodyIndex = 1 To sourcePart.Bodies.Count
        Set sourceBody = sourcePart.Bodies.Item(bodyIndex)

        ' Part.Bodies includes Bodies consumed by Boolean operations. Copy only
        ' active root result Bodies, otherwise operands would be counted twice.
        If BodyHasShape(sourceBody) Then
            If Not CBool(sourceBody.InBooleanOperation) Then
                If Not CBool(sourcePart.IsInactive(sourceBody)) Then
                    If copiedCount = 0 And Not BodyHasShape(firstBody) Then
                        Set targetBody = firstBody
                    Else
                        Set targetBody = targetPartDocument.Part.Bodies.Add
                    End If
                    targetBody.Name = "FLAT_RESULT_" & _
                                      Format$(copiedCount + failedCount + 1, "000") & _
                                      "__" & SafeFeatureName(sourceBody.Name)

                    pasteFailure = ""
                    If PasteObjectAsResult(sourcePartDocument, sourceBody, _
                                           targetPartDocument, targetBody, pasteFailure) Then
                        copiedCount = copiedCount + 1
                    Else
                        failedCount = failedCount + 1
                        mLastFlattenFailure = "Could not paste CATPart root result Body '" & _
                                              sourceBody.Name & "'. " & CStr(pasteFailure)
                        LogMessage "ERROR: " & mLastFlattenFailure
                        If Not targetBody Is firstBody Then
                            On Error Resume Next
                            DeleteObjectFromDocument targetPartDocument, targetBody
                            On Error GoTo 0
                        End If
                    End If
                End If
            End If
        End If
    Next bodyIndex
    If copiedCount = 0 And failedCount = 0 Then
        mLastFlattenFailure = "The CATPart contains no active root Body with a solid result."
    End If
End Sub

Private Sub FlattenProductLevel(ByVal containerProduct As Object, _
                                ByVal sourceProductDocument As Object, _
                                ByVal targetPartDocument As Object, _
                                ByRef copiedCount As Long, _
                                ByRef failedCount As Long, _
                                ByVal firstBody As Object)
    Dim i As Long
    Dim childProduct As Object
    Dim targetBody As Object
    Dim isPartInstance As Boolean
    Dim childCount As Long
    Dim pasteFailure As Variant

    childCount = containerProduct.Products.Count
    For i = 1 To childCount
        Set childProduct = containerProduct.Products.Item(i)
        SetStatus "Flattening component " & CStr(copiedCount + failedCount + 1) & _
                  ": " & childProduct.Name

        EnsureProductDesignMode childProduct, childProduct.Name

        isPartInstance = ProductInstanceReferencesPart(childProduct)
        If isPartInstance Then
            If copiedCount = 0 And BodyHasShape(firstBody) = False Then
                Set targetBody = firstBody
            Else
                Set targetBody = targetPartDocument.Part.Bodies.Add
            End If
            targetBody.Name = "FLAT_RESULT_" & Format$(copiedCount + failedCount + 1, "000") & _
                              "__" & SafeFeatureName(childProduct.Name)

            pasteFailure = ""
            If PasteObjectAsResult(sourceProductDocument, childProduct, _
                                   targetPartDocument, targetBody, pasteFailure) Then
                copiedCount = copiedCount + 1
            Else
                failedCount = failedCount + 1
                mLastFlattenFailure = _
                    "Could not paste product instance as an assembly-positioned result: " & _
                    childProduct.Name & ". " & CStr(pasteFailure)
                LogMessage "ERROR: " & mLastFlattenFailure
                If Not targetBody Is firstBody Then
                    On Error Resume Next
                    DeleteObjectFromDocument targetPartDocument, targetBody
                    On Error GoTo 0
                End If
            End If
        ElseIf ProductHasChildren(childProduct) Then
            FlattenProductLevel childProduct, sourceProductDocument, targetPartDocument, _
                                copiedCount, failedCount, firstBody
        ElseIf ProductHasMasterShape(childProduct) Then
            failedCount = failedCount + 1
            mLastFlattenFailure = "Unsupported non-CATPart shape representation: " & _
                                  childProduct.Name
            LogMessage "ERROR: " & mLastFlattenFailure
        Else
            ' A leaf with no master shape is an intentionally empty reference product.
            LogMessage "INFO: Skipped empty leaf product: " & childProduct.Name
        End If
    Next i
End Sub

Private Sub EnsureProductDesignMode(ByVal productObject As Object, ByVal productLabel As String)
    Dim modeErrorNumber As Long
    Dim modeErrorDescription As String

    On Error Resume Next
    Err.Clear
    productObject.ApplyWorkMode DESIGN_MODE
    modeErrorNumber = Err.Number
    modeErrorDescription = Err.Description
    Err.Clear
    On Error GoTo 0

    If modeErrorNumber <> 0 Then
        Err.Raise vbObjectError + 1721, "EnsureProductDesignMode", _
                  "Could not load '" & productLabel & "' in design mode (CATIA error " & _
                  CStr(modeErrorNumber) & "): " & modeErrorDescription
    End If
End Sub

Private Function ProductHasChildren(ByVal productObject As Object) As Boolean
    On Error GoTo NoChildren
    ProductHasChildren = (productObject.Products.Count > 0)
    Exit Function
NoChildren:
    ProductHasChildren = False
    Err.Clear
End Function

Private Function ProductHasMasterShape(ByVal productObject As Object) As Boolean
    On Error GoTo ShapeUnknown
    ProductHasMasterShape = CBool(productObject.HasAMasterShapeRepresentation)
    Exit Function
ShapeUnknown:
    ' Conservatively treat an unreadable leaf as geometry so the caller aborts.
    ProductHasMasterShape = True
    Err.Clear
End Function

Private Function ProductInstanceReferencesPart(ByVal productObject As Object) As Boolean
    Dim referenceDocument As Object

    On Error GoTo NotPart
    Set referenceDocument = productObject.ReferenceProduct.Parent
    ProductInstanceReferencesPart = (StrComp(TypeName(referenceDocument), "PartDocument", vbTextCompare) = 0)
    Exit Function
NotPart:
    ProductInstanceReferencesPart = False
    Err.Clear
End Function

Public Function PasteObjectAsResult(ByVal sourceDocument As Object, _
                                    ByVal sourceObject As Object, _
                                    ByVal targetPartDocument As Object, _
                                    ByVal targetBody As Object, _
                                    Optional ByRef failureText As Variant) As Boolean
    Dim sourceSelection As Object
    Dim targetSelection As Object
    Dim beforeShapeCount As Long
    Dim pasteErrorNumber As Long
    Dim pasteErrorDescription As String

    On Error GoTo PasteFailed
    If Not IsMissing(failureText) Then failureText = ""
    beforeShapeCount = targetBody.Shapes.Count

    sourceDocument.Activate
    Set sourceSelection = sourceDocument.Selection
    sourceSelection.Clear
    sourceSelection.Add sourceObject
    sourceSelection.Copy
    sourceSelection.Clear

    targetPartDocument.Activate
    targetPartDocument.Part.InWorkObject = targetBody
    Set targetSelection = targetPartDocument.Selection
    targetSelection.Clear
    targetSelection.Add targetBody
    targetSelection.PasteSpecial PASTE_AS_RESULT_WITHOUT_LINK
    targetSelection.Clear

    targetPartDocument.Part.UpdateObject targetBody
    PasteObjectAsResult = (targetBody.Shapes.Count > beforeShapeCount)
    If Not PasteObjectAsResult Then
        If Not IsMissing(failureText) Then
            failureText = "Paste Special completed but created no solid result in the target Body."
        End If
    End If
    Exit Function

PasteFailed:
    pasteErrorNumber = Err.Number
    pasteErrorDescription = Err.Description
    PasteObjectAsResult = False
    Err.Clear
    On Error Resume Next
    sourceDocument.Selection.Clear
    targetPartDocument.Selection.Clear
    On Error GoTo 0
    If Not IsMissing(failureText) Then
        failureText = "CATIA error " & CStr(pasteErrorNumber) & ": " & pasteErrorDescription
    End If
End Function

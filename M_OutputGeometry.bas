Attribute VB_Name = "M_OutputGeometry"
Option Explicit

Public Function BuildConstructionGridPart(ByRef gridSpec As TGridSpec) As Object
    Dim gridDocument As Object
    Dim partObject As Object
    Dim hybridFactory As Object
    Dim rootSet As Object
    Dim supportSet As Object
    Dim lineSet As Object
    Dim directionX As Object
    Dim directionY As Object
    Dim directionZ As Object
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim shapeCount As Long
    Dim totalLineCount As Long
    Dim xValue As Double
    Dim yValue As Double
    Dim zValue As Double

    SetStatus "Building the " & CStr(gridSpec.NX) & " x " & CStr(gridSpec.NY) & _
              " x " & CStr(gridSpec.NZ) & " construction lattice..."
    Set gridDocument = CATIA.Documents.Add("Part")
    Set partObject = gridDocument.Part
    gridDocument.Product.PartNumber = "GRID_CONSTRUCTION_" & gRunId
    Set hybridFactory = partObject.HybridShapeFactory

    Set rootSet = partObject.HybridBodies.Add
    rootSet.Name = "GRID_CONSTRUCTION__HIDE_THIS_GROUP"
    Set supportSet = rootSet.HybridBodies.Add
    supportSet.Name = "GRID_SUPPORT_POINTS__HIDDEN"
    Set lineSet = rootSet.HybridBodies.Add
    lineSet.Name = "GRID_LINES_" & CStr(gridSpec.NX) & "x" & CStr(gridSpec.NY) & "x" & CStr(gridSpec.NZ)
    totalLineCount = (gridSpec.NY + 1) * (gridSpec.NZ + 1) + _
                     (gridSpec.NX + 1) * (gridSpec.NZ + 1) + _
                     (gridSpec.NX + 1) * (gridSpec.NY + 1)

    Set directionX = hybridFactory.AddNewDirectionByCoord(1#, 0#, 0#)
    Set directionY = hybridFactory.AddNewDirectionByCoord(0#, 1#, 0#)
    Set directionZ = hybridFactory.AddNewDirectionByCoord(0#, 0#, 1#)

    ' X-parallel lines at every Y/Z boundary.
    For k = 0 To gridSpec.NZ
        zValue = gridSpec.Bounds.MinZ + CDbl(k) * gridSpec.DZ
        For j = 0 To gridSpec.NY
            yValue = gridSpec.Bounds.MinY + CDbl(j) * gridSpec.DY
            AddLatticeLine gridDocument, supportSet, lineSet, directionX, _
                           gridSpec.Bounds.MinX, yValue, zValue, _
                           gridSpec.Bounds.MaxX - gridSpec.Bounds.MinX, _
                           "GX_Y" & Format$(j, "00") & "_Z" & Format$(k, "00")
            shapeCount = shapeCount + 1
            GridProgress partObject, shapeCount, totalLineCount
        Next j
    Next k

    ' Y-parallel lines at every X/Z boundary.
    For k = 0 To gridSpec.NZ
        zValue = gridSpec.Bounds.MinZ + CDbl(k) * gridSpec.DZ
        For i = 0 To gridSpec.NX
            xValue = gridSpec.Bounds.MinX + CDbl(i) * gridSpec.DX
            AddLatticeLine gridDocument, supportSet, lineSet, directionY, _
                           xValue, gridSpec.Bounds.MinY, zValue, _
                           gridSpec.Bounds.MaxY - gridSpec.Bounds.MinY, _
                           "GY_X" & Format$(i, "00") & "_Z" & Format$(k, "00")
            shapeCount = shapeCount + 1
            GridProgress partObject, shapeCount, totalLineCount
        Next i
    Next k

    ' Z-parallel lines at every X/Y boundary.
    For j = 0 To gridSpec.NY
        yValue = gridSpec.Bounds.MinY + CDbl(j) * gridSpec.DY
        For i = 0 To gridSpec.NX
            xValue = gridSpec.Bounds.MinX + CDbl(i) * gridSpec.DX
            AddLatticeLine gridDocument, supportSet, lineSet, directionZ, _
                           xValue, yValue, gridSpec.Bounds.MinZ, _
                           gridSpec.Bounds.MaxZ - gridSpec.Bounds.MinZ, _
                           "GZ_X" & Format$(i, "00") & "_Y" & Format$(j, "00")
            shapeCount = shapeCount + 1
            GridProgress partObject, shapeCount, totalLineCount
        Next i
    Next j

    partObject.Update
    SetVisibility gridDocument, supportSet, catVisPropertyNoShowAttr
    SetVisibility gridDocument, lineSet, catVisPropertyShowAttr
    SetVisibility gridDocument, rootSet, catVisPropertyShowAttr
    SetVisibility gridDocument, partObject.Bodies.Item(1), catVisPropertyNoShowAttr
    StyleGridLines gridDocument, lineSet

    LogMessage "Construction grid: " & CStr(shapeCount) & " lattice line(s)."
    Set BuildConstructionGridPart = gridDocument
End Function

Private Sub AddLatticeLine(ByVal partDocument As Object, _
                           ByVal supportSet As Object, _
                           ByVal lineSet As Object, _
                           ByVal directionObject As Object, _
                           ByVal startX As Double, ByVal startY As Double, ByVal startZ As Double, _
                           ByVal lineLength As Double, _
                           ByVal lineName As String)
    Dim hybridFactory As Object
    Dim pointObject As Object
    Dim pointReference As Object
    Dim lineObject As Object

    Set hybridFactory = partDocument.Part.HybridShapeFactory
    Set pointObject = hybridFactory.AddNewPointCoord(startX, startY, startZ)
    pointObject.Name = "P_" & lineName
    supportSet.AppendHybridShape pointObject
    Set pointReference = partDocument.Part.CreateReferenceFromObject(pointObject)
    Set lineObject = hybridFactory.AddNewLinePtDir(pointReference, directionObject, 0#, lineLength, False)
    lineObject.Name = lineName
    lineSet.AppendHybridShape lineObject
End Sub

Private Sub GridProgress(ByVal partObject As Object, ByVal shapeCount As Long, ByVal totalLineCount As Long)
    If (shapeCount Mod GRID_UPDATE_INTERVAL) = 0 Then
        partObject.Update
        DoEvents
        SetStatus "Building construction lattice: " & Format$(shapeCount, "#,##0") & " of " & _
                  Format$(totalLineCount, "#,##0") & " lines"
    End If
End Sub

Private Sub StyleGridLines(ByVal gridDocument As Object, ByVal lineSet As Object)
    Dim selectionObject As Object
    Dim visualProperties As Object
    Dim styleErrorNumber As Long
    Dim styleErrorDescription As String

    On Error GoTo StyleFailed
    Set selectionObject = gridDocument.Selection
    selectionObject.Clear
    selectionObject.Add lineSet
    Set visualProperties = selectionObject.VisProperties
    visualProperties.SetRealColor COLOR_GRID_R, COLOR_GRID_G, COLOR_GRID_B, 1
    visualProperties.SetShow catVisPropertyShowAttr
    selectionObject.Clear
    Exit Sub

StyleFailed:
    styleErrorNumber = Err.Number
    styleErrorDescription = Err.Description
    Err.Clear
    On Error Resume Next
    gridDocument.Selection.Clear
    On Error GoTo 0
    Err.Raise vbObjectError + 1771, "StyleGridLines", _
              "Could not style the construction lattice (CATIA error " & _
              CStr(styleErrorNumber) & "): " & styleErrorDescription
End Sub

Public Function BuildCellSolidPart(ByRef gridSpec As TGridSpec, _
                                   ByRef flags() As Boolean, _
                                   ByVal resultPartNumber As String, _
                                   ByVal cellsAreAdded As Boolean) As Object
    Dim resultDocument As Object
    Dim partObject As Object
    Dim supportSet As Object
    Dim layerPlanes() As Object
    Dim planeObject As Object
    Dim cubeBody As Object
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim markedCount As Long
    Dim createdCount As Long
    Dim xMinimum As Double
    Dim yMinimum As Double
    Dim zMinimum As Double
    Dim cellName As String
    Dim promptResult As VbMsgBoxResult

    markedCount = CountMarkedCells(flags)
    If markedCount > OUTPUT_CELL_WARNING_COUNT Then
        promptResult = MsgBox( _
            Format$(markedCount, "#,##0") & " solid cell bodies will be created for " & _
            resultPartNumber & ". This can create a very large CATPart and take a long time." & _
            vbCrLf & vbCrLf & "Continue?", _
            vbYesNo + vbExclamation, "Grid Compare")
        If promptResult <> vbYes Then
            Err.Raise vbObjectError + 1770, "BuildCellSolidPart", _
                      "Solid-cell creation canceled by the user."
        End If
    End If

    SetStatus "Creating " & resultPartNumber & " solid cells..."
    Set resultDocument = CATIA.Documents.Add("Part")
    Set partObject = resultDocument.Part
    resultDocument.Product.PartNumber = resultPartNumber
    Set supportSet = partObject.HybridBodies.Add
    supportSet.Name = "CELL_SUPPORT_PLANES__HIDDEN"
    ReDim layerPlanes(0 To gridSpec.NZ - 1)

    For k = 0 To gridSpec.NZ - 1
        zMinimum = gridSpec.Bounds.MinZ + CDbl(k) * gridSpec.DZ
        For j = 0 To gridSpec.NY - 1
            yMinimum = gridSpec.Bounds.MinY + CDbl(j) * gridSpec.DY
            For i = 0 To gridSpec.NX - 1
                If flags(CellLinearIndex(gridSpec, i, j, k)) Then
                    If layerPlanes(k) Is Nothing Then
                        Set layerPlanes(k) = CreateOffsetXYPlane(resultDocument, supportSet, _
                            zMinimum, "CELL_PLANE_Z" & Format$(k, "00"))
                    End If
                    Set planeObject = layerPlanes(k)
                    xMinimum = gridSpec.Bounds.MinX + CDbl(i) * gridSpec.DX
                    cellName = "CELL_X" & Format$(i, "00") & _
                               "_Y" & Format$(j, "00") & "_Z" & Format$(k, "00")
                    Set cubeBody = CreateRectangularPrismBody(resultDocument, planeObject, _
                        xMinimum, yMinimum, xMinimum + gridSpec.DX, yMinimum + gridSpec.DY, _
                        gridSpec.DZ, cellName)
                    createdCount = createdCount + 1
                    If (createdCount Mod CUBE_UPDATE_INTERVAL) = 0 Then
                        DoEvents
                        SetStatus "Creating " & resultPartNumber & ": " & _
                                  Format$(createdCount, "#,##0") & " of " & _
                                  Format$(markedCount, "#,##0") & " solid cells"
                    End If
                End If
            Next i
        Next j
    Next k

    partObject.Update
    SetVisibility resultDocument, supportSet, catVisPropertyNoShowAttr
    SetVisibility resultDocument, partObject.Bodies.Item(1), catVisPropertyNoShowAttr
    If cellsAreAdded Then
        StyleCellBodies resultDocument, COLOR_BLUE_R, COLOR_BLUE_G, COLOR_BLUE_B
    Else
        StyleCellBodies resultDocument, COLOR_RED_R, COLOR_RED_G, COLOR_RED_B
    End If

    LogMessage resultPartNumber & ": created " & CStr(createdCount) & " solid cell body/bodies."
    Set BuildCellSolidPart = resultDocument
End Function

Private Sub StyleCellBodies(ByVal partDocument As Object, _
                            ByVal redValue As Long, ByVal greenValue As Long, ByVal blueValue As Long)
    Dim selectionObject As Object
    Dim visualProperties As Object
    Dim bodyIndex As Long
    Dim styleErrorNumber As Long
    Dim styleErrorDescription As String

    On Error GoTo StyleFailed
    Set selectionObject = partDocument.Selection
    selectionObject.Clear
    ' Item 1 is the intentionally empty default PartBody.
    For bodyIndex = 2 To partDocument.Part.Bodies.Count
        selectionObject.Add partDocument.Part.Bodies.Item(bodyIndex)
    Next bodyIndex
    If selectionObject.Count2 > 0 Then
        Set visualProperties = selectionObject.VisProperties
        visualProperties.SetRealColor redValue, greenValue, blueValue, 1
        visualProperties.SetRealOpacity HALF_OPACITY, 1
        visualProperties.SetShow catVisPropertyShowAttr
    End If
    selectionObject.Clear
    Exit Sub

StyleFailed:
    styleErrorNumber = Err.Number
    styleErrorDescription = Err.Description
    Err.Clear
    On Error Resume Next
    partDocument.Selection.Clear
    On Error GoTo 0
    Err.Raise vbObjectError + 1772, "StyleCellBodies", _
              "Could not style the solid comparison cells (CATIA error " & _
              CStr(styleErrorNumber) & "): " & styleErrorDescription
End Sub

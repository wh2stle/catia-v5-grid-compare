Attribute VB_Name = "M_Main"
Option Explicit

Public Sub CATMain()
    Dim originalPath As String
    Dim revisedPath As String
    Dim outputFolder As String
    Dim outputStem As String
    Dim masterPath As String
    Dim gridPath As String
    Dim addedCellsPath As String
    Dim removedCellsPath As String
    Dim addedViewPath As String
    Dim removedViewPath As String
    Dim reportPath As String
    Dim phaseName As String
    Dim preserveComputedDocuments As Boolean
    Dim originalOpenedByMacro As Boolean
    Dim revisedOpenedByMacro As Boolean

    Dim originalDocument As Object
    Dim revisedDocument As Object
    Dim masterDocument As Object
    Dim originalFlatDocument As Object
    Dim revisedFlatDocument As Object
    Dim addedDifferenceDocument As Object
    Dim removedDifferenceDocument As Object
    Dim addedDifferenceBody As Object
    Dim removedDifferenceBody As Object
    Dim gridDocument As Object
    Dim addedCellsDocument As Object
    Dim removedCellsDocument As Object
    Dim addedViewDocument As Object
    Dim removedViewDocument As Object
    Dim addedWindow As Object
    Dim removedWindow As Object

    Dim unionWithoutMargin As TBounds
    Dim gridSpec As TGridSpec
    Dim runStats As TRunStats
    Dim addedFlags() As Boolean
    Dim removedFlags() As Boolean
    Dim addedIsEmpty As Boolean
    Dim removedIsEmpty As Boolean
    Dim displayAlertsChanged As Boolean
    Dim originalDisplayFileAlerts As Boolean
    Dim totalCellCount As Long
    Dim failureNumber As Long
    Dim failureDescription As String

    On Error GoTo FatalError
    phaseName = "initialization"
    InitializeRunState
    ValidateConfiguration
    runStats.StartedAt = Now
    PreparePrivateTempFolder

    If CATIA.Windows.Count > 0 Or CATIA.Documents.Count > 0 Then
        MsgBox "Close every open CATIA document before starting Grid Compare." & vbCrLf & _
               "CATIA's tile command arranges all document windows, so an empty session is " & _
               "required to guarantee the requested two-panel result." & vbCrLf & vbCrLf & _
               "No file has been changed. Run the macro again after closing the windows.", _
               vbInformation, "Grid Compare"
        RemovePrivateTempFolder
        Exit Sub
    End If

    phaseName = "input selection"
    If Not SelectComparisonFiles(originalPath, revisedPath) Then
        RemovePrivateTempFolder
        Exit Sub
    End If
    LogMessage "Revision 1: " & originalPath
    LogMessage "Revised: " & revisedPath

    phaseName = "opening inputs"
    Set originalDocument = GetOrOpenCATIADocument(originalPath, originalOpenedByMacro)
    Set revisedDocument = GetOrOpenCATIADocument(revisedPath, revisedOpenedByMacro)
    ValidateInputDocument originalDocument, "revision 1"
    ValidateInputDocument revisedDocument, "revised"

    On Error Resume Next
    Err.Clear
    originalDisplayFileAlerts = CBool(CATIA.DisplayFileAlerts)
    If Err.Number = 0 Then
        CATIA.DisplayFileAlerts = False
        displayAlertsChanged = (Err.Number = 0)
    End If
    Err.Clear
    CATIA.RefreshDisplay = False
    Err.Clear
    On Error GoTo FatalError

    phaseName = "source product"
    Set masterDocument = CreateSourceMasterProduct(originalPath, revisedPath)

    phaseName = "flattening revision 1"
    Set originalFlatDocument = FlattenSourceToPart(originalDocument, _
        "FLAT_REVISION_1_" & gRunId, runStats.OriginalComponents, runStats.OriginalPasteFailures)
    phaseName = "flattening revised"
    Set revisedFlatDocument = FlattenSourceToPart(revisedDocument, _
        "FLAT_REVISED_" & gRunId, runStats.RevisedComponents, runStats.RevisedPasteFailures)

    If runStats.OriginalPasteFailures > 0 Or runStats.RevisedPasteFailures > 0 Then
        Err.Raise vbObjectError + 1780, "CATMain", _
                  "At least one CATPart instance could not be copied in assembly position. " & _
                  "The comparison was stopped to avoid an incomplete result."
    End If

    phaseName = "union bounding box"
    If Not ComputeUnionGrid(originalFlatDocument, revisedFlatDocument, _
                            unionWithoutMargin, gridSpec) Then
        Err.Raise vbObjectError + 1781, "CATMain", "Could not compute solid bounds for both inputs."
    End If

    phaseName = "added Boolean difference"
    Set addedDifferenceDocument = CreateDifferenceDocument(revisedFlatDocument, originalFlatDocument, _
        "ADDED_RAW_" & gRunId, addedDifferenceBody, runStats.AddedRawVolumeM3, addedIsEmpty)
    phaseName = "removed Boolean difference"
    Set removedDifferenceDocument = CreateDifferenceDocument(originalFlatDocument, revisedFlatDocument, _
        "REMOVED_RAW_" & gRunId, removedDifferenceBody, runStats.RemovedRawVolumeM3, removedIsEmpty)

    phaseName = "added grid scan"
    totalCellCount = CLng(CDbl(gridSpec.NX) * CDbl(gridSpec.NY) * CDbl(gridSpec.NZ))
    If addedIsEmpty Then
        ReDim addedFlags(0 To totalCellCount - 1)
    Else
        ScanDifferenceCells addedDifferenceDocument, addedDifferenceBody, gridSpec, _
                            addedFlags, runStats.AddedProbes, "Added"
    End If
    runStats.AddedCells = CountMarkedCells(addedFlags)

    phaseName = "removed grid scan"
    If removedIsEmpty Then
        ReDim removedFlags(0 To totalCellCount - 1)
    Else
        ScanDifferenceCells removedDifferenceDocument, removedDifferenceBody, gridSpec, _
                            removedFlags, runStats.RemovedProbes, "Removed"
    End If
    runStats.RemovedCells = CountMarkedCells(removedFlags)

    phaseName = "construction grid"
    Set gridDocument = BuildConstructionGridPart(gridSpec)
    phaseName = "added solid cells"
    Set addedCellsDocument = BuildCellSolidPart(gridSpec, addedFlags, _
        "ADDED_CELLS_" & gRunId, True)
    phaseName = "removed solid cells"
    Set removedCellsDocument = BuildCellSolidPart(gridSpec, removedFlags, _
        "REMOVED_CELLS_" & gRunId, False)

    ' The directory is deliberately requested only after computation is complete.
    phaseName = "save-folder selection"
    CATIA.RefreshDisplay = True
    outputFolder = BrowseForOutputFolder()
    If Len(outputFolder) = 0 Then
        preserveComputedDocuments = True
        Err.Raise vbObjectError + 1782, "CATMain", _
                  "No save folder was selected. Computed documents were left open and unsaved."
    End If
    CATIA.RefreshDisplay = False

    outputStem = MakeUniqueOutputStem(outputFolder)
    masterPath = JoinPath(outputFolder, outputStem & "_Source_Master.CATProduct")
    gridPath = JoinPath(outputFolder, outputStem & "_Grid_Construction.CATPart")
    addedCellsPath = JoinPath(outputFolder, outputStem & "_Added_Cells.CATPart")
    removedCellsPath = JoinPath(outputFolder, outputStem & "_Removed_Cells.CATPart")
    addedViewPath = JoinPath(outputFolder, outputStem & "_Added_View.CATProduct")
    removedViewPath = JoinPath(outputFolder, outputStem & "_Removed_View.CATProduct")
    reportPath = JoinPath(outputFolder, outputStem & "_Report.txt")

    phaseName = "saving source and geometry files"
    SaveDocumentAs masterDocument, masterPath
    SaveDocumentAs gridDocument, gridPath
    SaveDocumentAs addedCellsDocument, addedCellsPath
    SaveDocumentAs removedCellsDocument, removedCellsPath

    ' Close only macro-owned work documents. Inputs opened by the user are preserved.
    SafeCloseDocument addedDifferenceDocument
    Set addedDifferenceDocument = Nothing
    SafeCloseDocument removedDifferenceDocument
    Set removedDifferenceDocument = Nothing
    SafeCloseDocument originalFlatDocument
    Set originalFlatDocument = Nothing
    SafeCloseDocument revisedFlatDocument
    Set revisedFlatDocument = Nothing
    SafeCloseDocument gridDocument
    Set gridDocument = Nothing
    SafeCloseDocument addedCellsDocument
    Set addedCellsDocument = Nothing
    SafeCloseDocument removedCellsDocument
    Set removedCellsDocument = Nothing
    SafeCloseDocument masterDocument
    Set masterDocument = Nothing
    If originalOpenedByMacro Then
        SafeCloseDocument originalDocument
    End If
    If revisedOpenedByMacro Then
        SafeCloseDocument revisedDocument
    End If

    DoEvents
    If CATIA.Windows.Count <> 0 Then
        Err.Raise vbObjectError + 1784, "CATMain", _
                  "A temporary or input document window could not be closed. " & _
                  "The two-panel result was not created because CATIA would tile extra windows."
    End If

    phaseName = "removed view product"
    Set removedViewDocument = CreateResultViewProduct( _
        "REMOVED_VIEW_" & gRunId, originalPath, removedCellsPath, gridPath, False, removedWindow)
    SaveDocumentAs removedViewDocument, removedViewPath

    phaseName = "added view product"
    Set addedViewDocument = CreateResultViewProduct( _
        "ADDED_VIEW_" & gRunId, originalPath, addedCellsPath, gridPath, True, addedWindow)
    SaveDocumentAs addedViewDocument, addedViewPath

    runStats.FinishedAt = Now
    phaseName = "report"
    WriteUtf8TextFile reportPath, BuildRunReport(originalPath, revisedPath, _
        unionWithoutMargin, gridSpec, runStats, masterPath, gridPath, addedCellsPath, _
        removedCellsPath, addedViewPath, removedViewPath)
    LogMessage "Saved: " & reportPath
    RemovePrivateTempFolder

    On Error Resume Next
    CATIA.RefreshDisplay = True
    If displayAlertsChanged Then CATIA.DisplayFileAlerts = originalDisplayFileAlerts
    Err.Clear
    On Error GoTo 0

    phaseName = "camera synchronization"
    TileAndSynchronizeResultViews removedWindow, addedWindow
    removedViewDocument.Save
    addedViewDocument.Save
    MsgBox "Grid comparison complete." & vbCrLf & vbCrLf & _
           "Added cells: " & Format$(runStats.AddedCells, "#,##0") & vbCrLf & _
           "Removed cells: " & Format$(runStats.RemovedCells, "#,##0") & vbCrLf & vbCrLf & _
           "Files were saved in:" & vbCrLf & outputFolder, _
           vbInformation, "Grid Compare"
    Exit Sub

FatalError:
    failureNumber = Err.Number
    failureDescription = Err.Description
    Err.Clear

    On Error Resume Next
    CATIA.RefreshDisplay = True
    SetStatus "Grid Compare stopped during " & phaseName
    Err.Clear
    On Error GoTo 0

    If Not preserveComputedDocuments Then
        SafeCloseDocument addedViewDocument
        SafeCloseDocument removedViewDocument
        SafeCloseDocument addedDifferenceDocument
        SafeCloseDocument removedDifferenceDocument
        SafeCloseDocument originalFlatDocument
        SafeCloseDocument revisedFlatDocument
        SafeCloseDocument gridDocument
        SafeCloseDocument addedCellsDocument
        SafeCloseDocument removedCellsDocument
        SafeCloseDocument masterDocument
        If originalOpenedByMacro Then SafeCloseDocument originalDocument
        If revisedOpenedByMacro Then SafeCloseDocument revisedDocument
        ' Also catches a document created inside a function that failed before
        ' it could return its object reference to this procedure.
        CloseRunOwnedDocuments
    End If
    RemovePrivateTempFolder

    On Error Resume Next
    If displayAlertsChanged Then CATIA.DisplayFileAlerts = originalDisplayFileAlerts
    Err.Clear
    On Error GoTo 0

    MsgBox "Grid Compare stopped during " & phaseName & "." & vbCrLf & vbCrLf & _
           failureDescription & vbCrLf & "Error number: " & CStr(failureNumber), _
           vbCritical, "Grid Compare"
End Sub

Private Sub ValidateInputDocument(ByVal documentObject As Object, ByVal inputLabel As String)
    Dim documentType As String

    documentType = TypeName(documentObject)
    If StrComp(documentType, "PartDocument", vbTextCompare) <> 0 And _
       StrComp(documentType, "ProductDocument", vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 1783, "ValidateInputDocument", _
                  "The " & inputLabel & " input is not a CATPart or CATProduct."
    End If
End Sub

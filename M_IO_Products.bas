Attribute VB_Name = "M_IO_Products"
Option Explicit

Public Function SelectComparisonFiles(ByRef originalPath As String, ByRef revisedPath As String) As Boolean
    originalPath = CATIA.FileSelectionBox( _
        "Select revision 1 (original) CATPart or CATProduct", _
        "*.CATPart; *.CATProduct", CatFileSelectionModeOpen)
    If Len(originalPath) = 0 Then Exit Function
    If Not IsSupportedCATIAPath(originalPath) Then
        MsgBox "Revision 1 must be a CATPart or CATProduct.", vbExclamation, "Grid Compare"
        Exit Function
    End If

    revisedPath = CATIA.FileSelectionBox( _
        "Select the revised CATPart or CATProduct", _
        "*.CATPart; *.CATProduct", CatFileSelectionModeOpen)
    If Len(revisedPath) = 0 Then Exit Function
    If Not IsSupportedCATIAPath(revisedPath) Then
        MsgBox "The revised file must be a CATPart or CATProduct.", vbExclamation, "Grid Compare"
        Exit Function
    End If

    If StrComp(originalPath, revisedPath, vbTextCompare) = 0 Then
        MsgBox "Select two different files.", vbExclamation, "Grid Compare"
        Exit Function
    End If

    SelectComparisonFiles = True
End Function

Public Function GetOrOpenCATIADocument(ByVal fullPath As String, ByRef macroOpenedIt As Boolean) As Object
    Dim documentObject As Object
    Dim existingPath As String
    Dim savedStateKnown As Boolean
    Dim documentIsSaved As Boolean

    macroOpenedIt = False
    For Each documentObject In CATIA.Documents
        existingPath = ""
        On Error Resume Next
        Err.Clear
        existingPath = documentObject.FullName
        Err.Clear
        On Error GoTo 0
        If Len(existingPath) > 0 Then
            If StrComp(existingPath, fullPath, vbTextCompare) = 0 Then
                savedStateKnown = False
                On Error Resume Next
                Err.Clear
                documentIsSaved = CBool(documentObject.Saved)
                savedStateKnown = (Err.Number = 0)
                Err.Clear
                On Error GoTo 0
                If Not savedStateKnown Then
                    Err.Raise vbObjectError + 1711, "GetOrOpenCATIADocument", _
                              "Could not determine whether the already-open input is saved: " & fullPath
                End If
                If Not documentIsSaved Then
                    Err.Raise vbObjectError + 1712, "GetOrOpenCATIADocument", _
                              "The selected input has unsaved changes. Save it, then run the comparison again: " & _
                              fullPath
                End If
                Set GetOrOpenCATIADocument = documentObject
                Exit Function
            End If
        End If
    Next documentObject

    SetStatus "Opening " & FileLeaf(fullPath) & "..."
    Set GetOrOpenCATIADocument = CATIA.Documents.Open(fullPath)
    macroOpenedIt = True
End Function

Public Function BrowseForOutputFolder() As String
    Dim shellObject As Object
    Dim folderObject As Object
    Dim fso As Object
    Dim selectedPath As String

    Set shellObject = CreateObject("Shell.Application")
    Set folderObject = shellObject.BrowseForFolder(0, _
        "Select the folder for the CATIA grid-comparison result files", 1)
    If folderObject Is Nothing Then Exit Function

    On Error Resume Next
    selectedPath = folderObject.Self.Path
    On Error GoTo 0
    If Len(selectedPath) = 0 Then Exit Function

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(selectedPath) Then
        MsgBox "Please select a normal file-system folder.", vbExclamation, "Grid Compare"
        Exit Function
    End If
    BrowseForOutputFolder = selectedPath
End Function

Public Function MakeUniqueOutputStem(ByVal folderPath As String) As String
    Dim fso As Object
    Dim candidate As String
    Dim suffixNumber As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    candidate = "GridCompare_" & gRunId
    suffixNumber = 1

    Do
        If Not OutputStemExists(fso, folderPath, candidate) Then Exit Do
        suffixNumber = suffixNumber + 1
        candidate = "GridCompare_" & gRunId & "_" & Format$(suffixNumber, "00")
    Loop
    MakeUniqueOutputStem = candidate
End Function

Private Function OutputStemExists(ByVal fso As Object, _
                                  ByVal folderPath As String, _
                                  ByVal candidate As String) As Boolean
    Dim suffixes As Variant
    Dim suffixValue As Variant

    suffixes = Array( _
        "_Source_Master.CATProduct", _
        "_Grid_Construction.CATPart", _
        "_Added_Cells.CATPart", _
        "_Removed_Cells.CATPart", _
        "_Added_View.CATProduct", _
        "_Removed_View.CATProduct", _
        "_Report.txt")
    For Each suffixValue In suffixes
        If fso.FileExists(JoinPath(folderPath, candidate & CStr(suffixValue))) Then
            OutputStemExists = True
            Exit Function
        End If
    Next suffixValue
End Function

Public Function CreateSourceMasterProduct(ByVal originalPath As String, ByVal revisedPath As String) As Object
    Dim productDocument As Object
    Dim rootProduct As Object
    Dim originalComponent As Object
    Dim revisedComponent As Object

    SetStatus "Creating source comparison CATProduct..."
    Set productDocument = CATIA.Documents.Add("Product")
    Set rootProduct = productDocument.Product
    rootProduct.PartNumber = "GRID_COMPARE_SOURCE_" & gRunId
    rootProduct.Name = "GRID_COMPARE_SOURCE"

    Set originalComponent = AddOneComponentFile(rootProduct, originalPath)
    originalComponent.Name = "REVISION_1__" & SafeFeatureName(FileStem(originalPath))

    Set revisedComponent = AddOneComponentFile(rootProduct, revisedPath)
    revisedComponent.Name = "REVISED__" & SafeFeatureName(FileStem(revisedPath))

    productDocument.Activate
    Set CreateSourceMasterProduct = productDocument
End Function

Public Function AddOneComponentFile(ByVal rootProduct As Object, ByVal componentPath As String) As Object
    Dim componentFiles(0) As Variant
    Dim beforeCount As Long

    componentFiles(0) = componentPath
    beforeCount = rootProduct.Products.Count
    rootProduct.Products.AddComponentsFromFiles componentFiles, "All"
    If rootProduct.Products.Count <= beforeCount Then
        Err.Raise vbObjectError + 1710, "AddOneComponentFile", _
                  "CATIA did not add '" & componentPath & "' to the product."
    End If
    Set AddOneComponentFile = rootProduct.Products.Item(beforeCount + 1)
End Function

Public Function CreateResultViewProduct(ByVal viewName As String, _
                                        ByVal revisionOnePath As String, _
                                        ByVal overlayPath As String, _
                                        ByVal gridPath As String, _
                                        ByVal overlayIsAdded As Boolean, _
                                        ByRef viewWindow As Object) As Object
    Dim viewDocument As Object
    Dim rootProduct As Object
    Dim revisionComponent As Object
    Dim overlayComponent As Object
    Dim gridComponent As Object

    Set viewDocument = CATIA.Documents.Add("Product")
    Set rootProduct = viewDocument.Product
    rootProduct.PartNumber = viewName
    rootProduct.Name = viewName

    Set revisionComponent = AddOneComponentFile(rootProduct, revisionOnePath)
    revisionComponent.Name = "REVISION_1_GRAY"
    Set gridComponent = AddOneComponentFile(rootProduct, gridPath)
    gridComponent.Name = "GRID_CONSTRUCTION__HIDE_THIS_COMPONENT"
    Set overlayComponent = AddOneComponentFile(rootProduct, overlayPath)

    If overlayIsAdded Then
        overlayComponent.Name = "ADDED_CELLS_BLUE_50_PERCENT"
        SetAppearance viewDocument, overlayComponent, COLOR_BLUE_R, COLOR_BLUE_G, COLOR_BLUE_B, HALF_OPACITY
    Else
        overlayComponent.Name = "REMOVED_CELLS_RED_50_PERCENT"
        SetAppearance viewDocument, overlayComponent, COLOR_RED_R, COLOR_RED_G, COLOR_RED_B, HALF_OPACITY
    End If

    SetAppearance viewDocument, revisionComponent, COLOR_GRAY_R, COLOR_GRAY_G, COLOR_GRAY_B, OPAQUE
    SetAppearance viewDocument, gridComponent, COLOR_GRID_R, COLOR_GRID_G, COLOR_GRID_B, OPAQUE

    viewDocument.Activate
    Set viewWindow = CATIA.ActiveWindow
    Set CreateResultViewProduct = viewDocument
End Function

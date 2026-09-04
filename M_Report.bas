Attribute VB_Name = "M_Report"
Option Explicit

Public Function BuildRunReport(ByVal originalPath As String, _
                               ByVal revisedPath As String, _
                               ByRef unionWithoutMargin As TBounds, _
                               ByRef gridSpec As TGridSpec, _
                               ByRef runStats As TRunStats, _
                               ByVal masterPath As String, _
                               ByVal gridPath As String, _
                               ByVal addedCellsPath As String, _
                               ByVal removedCellsPath As String, _
                               ByVal addedViewPath As String, _
                               ByVal removedViewPath As String) As String
    Dim reportText As String

    reportText = gLogText & vbCrLf
    reportText = reportText & "INPUTS" & vbCrLf
    reportText = reportText & "Revision 1: " & originalPath & vbCrLf
    reportText = reportText & "Revised:    " & revisedPath & vbCrLf
    reportText = reportText & "Axis assumption: both inputs use the same model/assembly axis system." & vbCrLf
    reportText = reportText & "Engine: Part Design Booleans, GSD extrema, Product.Analyze volume; no DMU API." & _
                 vbCrLf & vbCrLf

    reportText = reportText & "GRID" & vbCrLf
    reportText = reportText & "Divisions: " & CStr(gridSpec.NX) & " x " & CStr(gridSpec.NY) & _
                 " x " & CStr(gridSpec.NZ) & vbCrLf
    reportText = reportText & "Margin: " & _
                 FormatInvariant(BOUNDS_MARGIN_FRACTION * 100#, 3) & _
                 "% of the union span on each side of every axis (" & _
                 FormatInvariant((1# + 2# * BOUNDS_MARGIN_FRACTION) * 100#, 3) & _
                 "% total span)." & vbCrLf
    reportText = reportText & BoundsText("Union before margin", unionWithoutMargin) & vbCrLf
    reportText = reportText & BoundsText("Final grid bounds", gridSpec.Bounds) & vbCrLf
    reportText = reportText & "Grid center mm: (" & _
                 FormatInvariant((gridSpec.Bounds.MinX + gridSpec.Bounds.MaxX) / 2#, 6) & ", " & _
                 FormatInvariant((gridSpec.Bounds.MinY + gridSpec.Bounds.MaxY) / 2#, 6) & ", " & _
                 FormatInvariant((gridSpec.Bounds.MinZ + gridSpec.Bounds.MaxZ) / 2#, 6) & ")" & vbCrLf
    reportText = reportText & "Center boundary indices: X=" & CStr(gridSpec.NX \ 2) & _
                 ", Y=" & CStr(gridSpec.NY \ 2) & ", Z=" & CStr(gridSpec.NZ \ 2) & "." & vbCrLf
    reportText = reportText & "Cell size mm: " & FormatInvariant(gridSpec.DX, 6) & " x " & _
                 FormatInvariant(gridSpec.DY, 6) & " x " & FormatInvariant(gridSpec.DZ, 6) & vbCrLf & vbCrLf

    reportText = reportText & "RESULTS" & vbCrLf
    reportText = reportText & "Revision 1 flattened CATPart items: " & _
                 CStr(runStats.OriginalComponents) & vbCrLf
    reportText = reportText & "Revised flattened CATPart items:    " & _
                 CStr(runStats.RevisedComponents) & vbCrLf
    reportText = reportText & "Paste failures: revision 1=" & _
                 CStr(runStats.OriginalPasteFailures) & ", revised=" & _
                 CStr(runStats.RevisedPasteFailures) & vbCrLf
    reportText = reportText & "Added raw Boolean volume m^3:   " & _
                 FormatInvariant(runStats.AddedRawVolumeM3, 12) & vbCrLf
    reportText = reportText & "Removed raw Boolean volume m^3: " & _
                 FormatInvariant(runStats.RemovedRawVolumeM3, 12) & vbCrLf
    reportText = reportText & "Added marked cells:   " & CStr(runStats.AddedCells) & vbCrLf
    reportText = reportText & "Removed marked cells: " & CStr(runStats.RemovedCells) & vbCrLf
    reportText = reportText & "Added Boolean probes:   " & CStr(runStats.AddedProbes) & vbCrLf
    reportText = reportText & "Removed Boolean probes: " & CStr(runStats.RemovedProbes) & vbCrLf & vbCrLf

    reportText = reportText & "0.1 MM FILTER - IMPORTANT" & vbCrLf
    reportText = reportText & "CATIA V5 R27 VBA exposes exact Boolean geometry but not a general minimum-thickness " & _
                 "operator. This macro therefore rejects an entire raw added/removed result below " & _
                 FormatInvariant(MIN_CHANGE_VOLUME_M3, 12) & " m^3 (the volume of a " & _
                 FormatInvariant(MIN_CHANGE_MM, 3) & " mm cube). After that run-level filter passes, " & _
                 "every cell with a positive Boolean-intersection volume is reported. CATIA's Boolean " & _
                 "kernel is the geometric tolerance gate. This " & _
                 "preserves the requested any-portion cell rule, but it is not a guaranteed 0.1 mm thickness " & _
                 "test: a broad feature thinner than 0.1 mm can have more volume and be reported." & vbCrLf & vbCrLf

    reportText = reportText & "OUTPUT FILES" & vbCrLf
    reportText = reportText & "Source master: " & masterPath & vbCrLf
    reportText = reportText & "Construction grid: " & gridPath & vbCrLf
    reportText = reportText & "Added cells: " & addedCellsPath & vbCrLf
    reportText = reportText & "Removed cells: " & removedCellsPath & vbCrLf
    reportText = reportText & "Added view: " & addedViewPath & vbCrLf
    reportText = reportText & "Removed view: " & removedViewPath & vbCrLf & vbCrLf

    reportText = reportText & "DISPLAY" & vbCrLf
    reportText = reportText & "Both view products use revision 1 in opaque gray. Added cells are blue at opacity " & _
                 CStr(HALF_OPACITY) & "/255; removed cells are red at opacity " & _
                 CStr(HALF_OPACITY) & "/255. The construction grid is visible initially. Hide the " & _
                 "GRID_CONSTRUCTION component, or its GRID_CONSTRUCTION__HIDE_THIS_GROUP set, to hide it." & vbCrLf
    reportText = reportText & "The live camera link runs until Esc is pressed." & vbCrLf & vbCrLf

    reportText = reportText & "TEMPORARY DATA" & vbCrLf
    reportText = reportText & "Only private folders beginning with " & TEMP_PREFIX & _
                 " are purged. CATIA's shared/global cache is never deleted." & vbCrLf
    reportText = reportText & "Finished: " & CStr(runStats.FinishedAt) & vbCrLf
    BuildRunReport = reportText
End Function

Private Function BoundsText(ByVal labelText As String, ByRef boundsValue As TBounds) As String
    BoundsText = labelText & " mm: X[" & FormatInvariant(boundsValue.MinX, 6) & ", " & _
                 FormatInvariant(boundsValue.MaxX, 6) & "] Y[" & _
                 FormatInvariant(boundsValue.MinY, 6) & ", " & _
                 FormatInvariant(boundsValue.MaxY, 6) & "] Z[" & _
                 FormatInvariant(boundsValue.MinZ, 6) & ", " & _
                 FormatInvariant(boundsValue.MaxZ, 6) & "]"
End Function

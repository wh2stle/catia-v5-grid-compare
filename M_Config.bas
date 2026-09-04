Attribute VB_Name = "M_Config"
Option Explicit

' Requested defaults. Change only these values to tune a run.
Public Const GRID_DIVISIONS_X As Long = 50
Public Const GRID_DIVISIONS_Y As Long = 50
Public Const GRID_DIVISIONS_Z As Long = 50
Public Const BOUNDS_MARGIN_FRACTION As Double = 0.05
Public Const MIN_CHANGE_MM As Double = 0.1

' A 0.1 mm cube is 0.001 mm^3, which is 1E-12 m^3.  V5 Automation does
' not expose a topology-independent minimum-thickness operator, so this is
' used as the documented Boolean-volume noise floor.
Public Const MIN_CHANGE_VOLUME_M3 As Double = _
    (MIN_CHANGE_MM * MIN_CHANGE_MM * MIN_CHANGE_MM) / 1000000000#
' A successful cell intersection is retained whenever Analyze.Volume is positive.
' CATIA's Boolean kernel remains the geometric tolerance gate.
Public Const EMPTY_VOLUME_EPS_M3 As Double = 0#

Public Const OUTPUT_CELL_WARNING_COUNT As Long = 25000
Public Const PROGRESS_INTERVAL As Long = 25
Public Const GRID_UPDATE_INTERVAL As Long = 250
Public Const CUBE_UPDATE_INTERVAL As Long = 100

' CATIA visual values: opacity 0 is transparent, 255 is opaque.
Public Const HALF_OPACITY As Long = 128
Public Const OPAQUE As Long = 255
Public Const COLOR_RED_R As Long = 255
Public Const COLOR_RED_G As Long = 0
Public Const COLOR_RED_B As Long = 0
Public Const COLOR_BLUE_R As Long = 0
Public Const COLOR_BLUE_G As Long = 0
Public Const COLOR_BLUE_B As Long = 255
Public Const COLOR_GRAY_R As Long = 165
Public Const COLOR_GRAY_G As Long = 165
Public Const COLOR_GRAY_B As Long = 165
Public Const COLOR_GRID_R As Long = 95
Public Const COLOR_GRID_G As Long = 95
Public Const COLOR_GRID_B As Long = 95

Public Const PASTE_AS_RESULT_WITHOUT_LINK As String = "CATPrtResultWithOutLink"
Public Const TEMP_PREFIX As String = "CATIA_GridCompare_Work_"

Public Sub ValidateConfiguration()
    If GRID_DIVISIONS_X <= 0 Or GRID_DIVISIONS_Y <= 0 Or GRID_DIVISIONS_Z <= 0 Then
        Err.Raise vbObjectError + 1690, "ValidateConfiguration", _
                  "Every grid-division count must be positive."
    End If
    If (GRID_DIVISIONS_X Mod 2) <> 0 Or (GRID_DIVISIONS_Y Mod 2) <> 0 Or _
       (GRID_DIVISIONS_Z Mod 2) <> 0 Then
        Err.Raise vbObjectError + 1691, "ValidateConfiguration", _
                  "Every grid-division count must be even so the center is a grid boundary."
    End If
    If CDbl(GRID_DIVISIONS_X) * CDbl(GRID_DIVISIONS_Y) * _
       CDbl(GRID_DIVISIONS_Z) > 2147483647# Then
        Err.Raise vbObjectError + 1692, "ValidateConfiguration", _
                  "The total cell count exceeds the VBA Long index limit."
    End If
    If BOUNDS_MARGIN_FRACTION < 0# Then
        Err.Raise vbObjectError + 1693, "ValidateConfiguration", _
                  "The bounding margin cannot be negative."
    End If
    If MIN_CHANGE_MM <= 0# Then
        Err.Raise vbObjectError + 1694, "ValidateConfiguration", _
                  "The minimum-change setting must be positive."
    End If
End Sub

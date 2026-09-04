Attribute VB_Name = "M_Types"
Option Explicit

' Axis-aligned bounds, expressed in the common model axis system, in millimetres.
Public Type TBounds
    MinX As Double
    MinY As Double
    MinZ As Double
    MaxX As Double
    MaxY As Double
    MaxZ As Double
End Type

Public Type TGridSpec
    Bounds As TBounds
    NX As Long
    NY As Long
    NZ As Long
    DX As Double
    DY As Double
    DZ As Double
End Type

Public Type TRunStats
    OriginalComponents As Long
    RevisedComponents As Long
    OriginalPasteFailures As Long
    RevisedPasteFailures As Long
    AddedCells As Long
    RemovedCells As Long
    AddedProbes As Long
    RemovedProbes As Long
    AddedRawVolumeM3 As Double
    RemovedRawVolumeM3 As Double
    StartedAt As Date
    FinishedAt As Date
End Type


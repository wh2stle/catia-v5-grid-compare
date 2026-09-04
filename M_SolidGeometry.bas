Attribute VB_Name = "M_SolidGeometry"
Option Explicit

Public Function CreateOffsetXYPlane(ByVal partDocument As Object, _
                                    ByVal supportSet As Object, _
                                    ByVal zOffsetMM As Double, _
                                    ByVal planeName As String) As Object
    Dim partObject As Object
    Dim hybridFactory As Object
    Dim xyReference As Object
    Dim planeObject As Object

    Set partObject = partDocument.Part
    Set hybridFactory = partObject.HybridShapeFactory
    Set xyReference = partObject.CreateReferenceFromObject(partObject.OriginElements.PlaneXY)
    Set planeObject = hybridFactory.AddNewPlaneOffset(xyReference, zOffsetMM, False)
    planeObject.Name = planeName
    supportSet.AppendHybridShape planeObject
    partObject.UpdateObject planeObject
    Set CreateOffsetXYPlane = planeObject
End Function

Public Function CreateRectangularPrismBody(ByVal partDocument As Object, _
                                           ByVal planeObject As Object, _
                                           ByVal xMinimumMM As Double, _
                                           ByVal yMinimumMM As Double, _
                                           ByVal xMaximumMM As Double, _
                                           ByVal yMaximumMM As Double, _
                                           ByVal heightMM As Double, _
                                           ByVal bodyName As String) As Object
    Dim partObject As Object
    Dim bodyObject As Object
    Dim sketchObject As Object
    Dim sketchFactory As Object
    Dim planeReference As Object
    Dim shapeFactory As Object
    Dim padObject As Object

    If xMaximumMM <= xMinimumMM Or yMaximumMM <= yMinimumMM Or heightMM <= 0# Then
        Err.Raise vbObjectError + 1750, "CreateRectangularPrismBody", _
                  "A prism dimension is zero or negative."
    End If

    Set partObject = partDocument.Part
    Set bodyObject = partObject.Bodies.Add
    bodyObject.Name = bodyName
    partObject.InWorkObject = bodyObject

    Set planeReference = partObject.CreateReferenceFromObject(planeObject)
    Set sketchObject = bodyObject.Sketches.Add(planeReference)
    sketchObject.Name = bodyName & "_PROFILE"
    Set sketchFactory = sketchObject.OpenEdition
    sketchFactory.CreateLine xMinimumMM, yMinimumMM, xMaximumMM, yMinimumMM
    sketchFactory.CreateLine xMaximumMM, yMinimumMM, xMaximumMM, yMaximumMM
    sketchFactory.CreateLine xMaximumMM, yMaximumMM, xMinimumMM, yMaximumMM
    sketchFactory.CreateLine xMinimumMM, yMaximumMM, xMinimumMM, yMinimumMM
    sketchObject.CloseEdition
    partObject.UpdateObject sketchObject

    Set shapeFactory = partObject.ShapeFactory
    partObject.InWorkObject = bodyObject
    Set padObject = shapeFactory.AddNewPad(sketchObject, heightMM)
    padObject.Name = bodyName & "_SOLID"
    partObject.UpdateObject padObject
    Set CreateRectangularPrismBody = bodyObject
End Function


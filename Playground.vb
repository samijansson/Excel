Sub CreateWorksheetIfNotExists(name As String)
    If Not WorksheetExists(name) Then
        Worksheets.Add.name = name
    End If
End Sub

Function WorksheetExists(name As String) As Boolean
    Dim wksht As Worksheet
    WorksheetExists = False
    For Each wksht In Worksheets
        If wksht.name = name Then
            WorksheetExists = True
            Exit For
        End If
    Next wksht
End Function



Sub PositionShapesInsideShape()
    ' Create the outer rectangle shape
    Dim rectangle As Shape
    Set rectangle = ActiveSheet.Shapes.AddShape(msoShapeRectangle, 50, 50, 200, 100)
    rectangle.Fill.ForeColor.RGB = RGB(255, 255, 255) ' Set the fill color to white

    ' Determine the position and size of the circle shapes relative to the rectangle
    Dim circle_diameter As Double
    circle_diameter = 50
    Dim circle_spacing As Double
    circle_spacing = 10
    Dim circle_left As Double
    circle_left = rectangle.left + (rectangle.width / 2) - ((circle_diameter + circle_spacing) / 2)
    Dim circle_top As Double
    circle_top = rectangle.top + (rectangle.height / 2) - (circle_diameter / 2)

    ' Create the inner circle shapes
    Dim circle1 As Shape
    Set circle1 = ActiveSheet.Shapes.AddShape(msoShapeOval, circle_left, circle_top, circle_diameter, circle_diameter)
    circle1.Fill.ForeColor.RGB = RGB(255, 0, 0) ' Set the fill color to red
    Dim circle2 As Shape
    Set circle2 = ActiveSheet.Shapes.AddShape(msoShapeOval, circle_left + circle_diameter + circle_spacing, circle_top, circle_diameter, circle_diameter)
    circle2.Fill.ForeColor.RGB = RGB(0, 0, 255) ' Set the fill color to blue
End Sub

Sub PositionRectanglesInTriangle()

    ' Set up triangle and rectangle shapes
    Dim triangle As Shape
    Dim rectangle1 As Shape
    Dim rectangle2 As Shape
    
    Set triangle = ActiveSheet.Shapes.AddShape(msoShapeRightTriangle, 100, 100, 200, 200)
    Set rectangle1 = ActiveSheet.Shapes.AddShape(msoShapeRectangle, 0, 0, 50, 50)
   Set rectangle2 = ActiveSheet.Shapes.AddShape(msoShapeRectangle, 0, 0, 50, 50)
    
    ' Position first rectangle
    rectangle1.left = triangle.left + 2
    rectangle1.top = triangle.top + triangle.left - 2
    
    ' Position second rectangle
    rectangle2.left = triangle.left + (triangle.width / 2) - (rectangle2.width / 2)
    rectangle2.top = triangle.top + triangle.height - rectangle2.height
    
End Sub

Sub CreateRightTriangleByArea(area As Double)
    Dim base As Double
    Dim height As Double
    
    Dim triangleArea As Double
    
    
    For base = 1 To 100 ' loop through possible values of the base
        For height = 1 To 100 ' loop through possible values of the height
            triangleArea = 0.5 * base * height ' calculate the area of the right triangle
            If triangleArea = area Then ' check if the area matches our desired area
                ' create a shape on the Excel worksheet using VBA
                Dim shp As Shape
                Set shp = ActiveSheet.Shapes.AddShape(msoShapeRightTriangle, 100, 100, base, height)
                Exit For ' exit the subroutine once we find the correct values for the base and height
            End If
        Next height
    Next base
End Sub

Function DetermineObjectLocation(myShape As Variant)
    Dim ShapeObjecCoord As New Dictionary
    ShapeObjecCoord.Add "Top", myShape.top
    ShapeObjecCoord.Add "Left", myShape.left
    ShapeObjecCoord.Add "Height", myShape.height
    ShapeObjecCoord.Add "Width", myShape.width
    ShapeObjecCoord.Add "Rotation", myShape.Rotation
    Set DetermineObjectLocation = ShapeObjecCoord
End Function

Function CalculateTriangleArea(base As Double, height As Double) As Double
    CalculateTriangleArea = 0.5 * base * height ' calculate the area using the formula
End Function

Function CalculateSphereArea(radius As Double)
    Const pi As Double = 3.14159 ' define pi as a constant
    CalculateSphereArea = 4 * pi * radius ^ 2 ' calculate the area using the formula
End Function



Function CalculateOvalArea(ByVal width As Double, ByVal height As Double) As Double
    Dim majorAxis As Double
    Dim minorAxis As Double
    Dim area As Double
    
    majorAxis = GetOvalMajorAxisLengths
    minorAxis = GetOvalMinorAxisLengths
    
    area = WorksheetFunction.pi * (majorAxis / 2) * (minorAxis / 2) ' calculate the area using the formula
    CalculateOvalArea = area
End Function

Function GetOvalMajorAxisLengths(ByVal width As Double, ByVal height As Double) As Double
    Dim majorAxis As Double
   'Calculate the lengths of the axes using the formula
    majorAxis = Sqr((width / 2) ^ 2 + (height / 2) ^ 2) * 2
    
    'Return the lengths of the axes as a formatted string
    GetOvalMajorAxisLengths = majorAxis
End Function

Function GetOvalMinorAxisLengths(ByVal width As Double, ByVal height As Double) As Double
    Dim minorAxis As Double
    'Calculate the lengths of the axes using the formula
    minorAxis = Sqr((width / 2) ^ 2 + (height / 2) ^ 2 / 4) * 2
    
    'Return the lengths of the axes as a formatted string
    GetOvalMinorAxisLengths = minorAxis
End Function

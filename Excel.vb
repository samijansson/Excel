Const TEAM_MAP_SHEET_NAME As String = "Team Map"
Const SYSTEM_COVERAGE_SHEET_NAME As String = "System Coverage"
Const DATA_SHEET_NAME As String = "Data"

'determine the coordinates of the three vertices of a right triangle
Function GetTriangleVertices(triangle As Shape) As Variant
    ' Get the coordinates of the triangle's top left corner
    Dim left As Double
    Dim top As Double
    left = triangle.left
    top = triangle.top
    
    ' Get the height and width of the triangle
    Dim height As Double
    Dim width As Double
    height = triangle.height
    width = triangle.width
    
    ' Determine the coordinates of the vertices
    Dim vertices(1 To 3, 1 To 2) As Double
    vertices(1, 1) = left ' Vertex 1 (top left corner)
    vertices(1, 2) = top
    vertices(2, 1) = left ' Vertex 2 (bottom left corner)
    vertices(2, 2) = top + height
    vertices(3, 1) = left + width ' Vertex 3 (top right corner)
    vertices(3, 2) = top
    
    GetTriangleVertices = vertices
End Function

'calculate the angle between the hypotenuse and the vertical axis, 
'and converts the result from radians to degrees.
Function GetHypotenuseAngleVertical(triangle As Shape) As Double
    ' Get the coordinates of the triangle's vertices
    Dim vertices As Variant
    vertices = GetTriangleVertices(triangle)
    Dim x1 As Double
    Dim y1 As Double
    Dim x2 As Double
    Dim y2 As Double
    x1 = vertices(1, 1)
    y1 = vertices(1, 2)
    x2 = vertices(3, 1)
    y2 = vertices(3, 2)
    
    ' Calculate the length of the hypotenuse and the height of the triangle
    Dim hypotenuse As Double
    Dim height As Double
    hypotenuse = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
    height = Abs(y2 - y1)
    
    ' Calculate the angle between the hypotenuse and the vertical axis
    Dim angle As Double
    angle = Atn(height / (hypotenuse / 2)) * 180 / WorksheetFunction.Pi
    
    GetHypotenuseAngleVertical = angle
End Function

'Calculate the angle between the hypotenuse and the horizontal axis,
Function GetHypotenuseAngleHorizontal(triangle As Shape, isHorizontal As Boolean) As Double
    ' Get the coordinates of the triangle's vertices
    Dim vertices As Variant
    vertices = GetTriangleVertices(triangle)
    Dim x1 As Double
    Dim y1 As Double
    Dim x2 As Double
    Dim y2 As Double
    x1 = vertices(1, 1)
    y1 = vertices(1, 2)
    x2 = vertices(3, 1)
    y2 = vertices(3, 2)
    
    ' Calculate the length of the hypotenuse and the width or height of the triangle
    Dim hypotenuse As Double
    Dim distance As Double
    hypotenuse = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
    If isHorizontal Then
        distance = Abs(x2 - x1)
    Else
        distance = Abs(y2 - y1)
    End If
    
    ' Calculate the angle between the hypotenuse and the horizontal or vertical axis
    Dim angle As Double
    angle = Atn(distance / (hypotenuse / 2)) * 180 / WorksheetFunction.Pi
    
    GetHypotenuseAngleHorizontal = angle
End Function

'decides on the positioning of the rectangle within the triangle
Function GetRectanglePositionInTriangle(rectangleWidth As Double, rectangleHeight As Double, _
                                         triangleX1 As Double, triangleY1 As Double, _
                                         triangleX2 As Double, triangleY2 As Double, _
                                         triangleX3 As Double, triangleY3 As Double) As String

    ' Calculate the center point of the triangle.
    Dim centerX As Double, centerY As Double
    centerX = (triangleX1 + triangleX2 + triangleX3) / 3
    centerY = (triangleY1 + triangleY2 + triangleY3) / 3
    
    ' Calculate the distance between the center point and each vertex of the triangle.
    Dim distance1 As Double, distance2 As Double, distance3 As Double
    distance1 = WorksheetFunction.Sqrt((triangleX1 - centerX) ^ 2 + (triangleY1 - centerY) ^ 2)
    distance2 = WorksheetFunction.Sqrt((triangleX2 - centerX) ^ 2 + (triangleY2 - centerY) ^ 2)
    distance3 = WorksheetFunction.Sqrt((triangleX3 - centerX) ^ 2 + (triangleY3 - centerY) ^ 2)
    
    ' Determine the relative position of the rectangle within the triangle based on the position of
    ' each vertex with respect to the center point.
    If distance1 >= distance2 And distance1 >= distance3 Then
        If triangleX1 <= centerX Then
            GetRectanglePositionInTriangle = "left"
        Else
            GetRectanglePositionInTriangle = "right"
        End If
    ElseIf distance2 >= distance1 And distance2 >= distance3 Then
        If triangleX2 <= centerX Then
            GetRectanglePositionInTriangle = "left"
        Else
            GetRectanglePositionInTriangle = "right"
        End If
    Else
        If triangleX3 <= centerX Then
            GetRectanglePositionInTriangle = "left"
        Else
            GetRectanglePositionInTriangle = "right"
        End If
    End If
    
    ' If the rectangle is positioned on the left or right side of the triangle, also determine
    ' whether it should be positioned at the top, bottom, or center of the side.
    If GetRectanglePositionInTriangle = "left" Or GetRectanglePositionInTriangle = "right" Then
        Dim triangleHeight As Double
        triangleHeight = WorksheetFunction.Max(triangleY1, triangleY2, triangleY3) - WorksheetFunction.Min(triangleY1, triangleY2, triangleY3)
        Dim rectanglePositionFactor As Double
        rectanglePositionFactor = (centerY - WorksheetFunction.Min(triangleY1, triangleY2, triangleY3)) / triangleHeight
        If rectanglePositionFactor <= 0.25 Then
            GetRectanglePositionInTriangle = GetRectanglePositionInTriangle & " top"
        ElseIf rectanglePositionFactor >= 0.75 Then
            GetRectanglePositionInTriangle = GetRectanglePositionInTriangle & " bottom"
        Else
            GetRectanglePositionInTriangle = GetRectanglePositionInTriangle & " center"
        End If
    End If
    
    ' If the rectangle is positioned at the top or bottom of the triangle, also determine
    ' whether it should be positioned on the left, right, or center of the side.
    If GetRectanglePositionInTriangle = "top" Or GetRectanglePositionInTriangle = "bottom" Then
        Dim triangleWidth As Double
        triangleWidth = WorksheetFunction.Max(triangleX1, triangleX2, triangleX3) - WorksheetFunction.Min(triangleX1, triangleX2,
        'triangleWidth = WorksheetFunction.Sqrt((triangleX1 - triangleX2) ^ 2 + (triangleY1 - triangleY2) ^ 2)
        Dim rectanglePositionFactor As Double
        rectanglePositionFactor = (centerX - WorksheetFunction.Min(triangleX1, triangleX2, triangleX3)) / triangleWidth
        If rectanglePositionFactor <= 0.25 Then
            GetRectanglePositionInTriangle = "left " & GetRectanglePositionInTriangle
        ElseIf rectanglePositionFactor >= 0.75 Then
            GetRectanglePositionInTriangle = "right " & GetRectanglePositionInTriangle
        Else
            GetRectanglePositionInTriangle = "center " & GetRectanglePositionInTriangle
        End If
    End If
    GetRectanglePositionInTriangle = Trim(GetRectanglePositionInTriangle)
End Function

'calculate the distance from the vertex of the triangle to the rectangl
Function GetDistanceFromVertexToRectangle(vertexX As Double, vertexY As Double, rectangleLeft As Double, rectangleTop As Double, rectangleWidth As Double, rectangleHeight As Double) As Double
    ' Calculate the center point of the rectangle
    Dim rectangleCenterX As Double
    Dim rectangleCenterY As Double
    rectangleCenterX = rectangleLeft + (rectangleWidth / 2)
    rectangleCenterY = rectangleTop + (rectangleHeight / 2)
    
    ' Calculate the distance from the vertex to the center of the rectangle
    Dim distanceToCenter As Double
    distanceToCenter = ((rectangleCenterX - vertexX) ^ 2 + (rectangleCenterY - vertexY) ^ 2) ^ 0.5
    
    ' Calculate the distance from the vertex to the edge of the rectangle
    Dim distanceToEdge As Double
    If rectangleCenterX < vertexX Then
        distanceToEdge = rectangleCenterX - rectangleLeft
    Else
        distanceToEdge = rectangleLeft + rectangleWidth - rectangleCenterX
    End If
    
    ' Calculate the distance from the vertex to the rectangle
    Dim distanceToRectangle As Double
    distanceToRectangle = (distanceToCenter ^ 2 - distanceToEdge ^ 2) ^ 0.5
    
    GetDistanceFromVertexToRectangle = distanceToRectangle
End Function

Sub CreateSystemDiagramTemplate()
    Dim triangleHeight As New Dictionary
    Dim triangleWidth As New Dictionary

    Dim wksht As Worksheet
    Set wksht = Worksheets(TEAM_MAP_SHEET_NAME)
    
    Dim ListOfSystemInProjectDictKeys As New Dictionary
    Set ListOfSystemInProjectDictKeys = TeamMapListOfSystemsInAProject
    Dim projectKey As Variant
    Dim systemCollection As Collection
    For Each projectKey In ListOfSystemInProjectDictKeys.Keys
        'Debug.Print "Project: " & projectKey
		Set systemCollection = ListOfSystemInProjectDictKeys(projectKey)
		Dim TotalHeightByProject As Long
		Dim TotalWeidthByProject As Long
		TotalHeightByProject = 0
		TotalWeidthByProject = 0
        
        Dim measurements As Variant
        measurements = GetShapeMeasurements()
        Dim key As Variant
        Dim dict As Object
        Set dict = measurements(1)
    
        Dim systemName As Variant
        For Each systemName In systemCollection
            ' Iterate through the dictionary and print the measurements for each shape
            For Each key In dict.Keys
                Dim index As Long
                index = dict(key)
                If key = systemName Then
                    TotalHeightByProject = TotalHeightByProject + measurements(0)(1, index)
                    TotalWeidthByProject = TotalWeidthByProject + measurements(0)(2, index)
                    Exit For
                End If
            Next key
        Next systemName
        
        If Not triangleHeight.Exists(projectKey) Then
            triangleHeight.Add projectKey, TotalHeightByProject
        End If
		
        If Not triangleWidth.Exists(projectKey) Then
            triangleWidth.Add projectKey, TotalWeidthByProject
        End If		
    Next projectKey

    Dim MaxTriangleHeight As Long
    MaxTriangleHeight = GetMaxItemValue(triangleHeight) / 3
    
    Dim MaxTriangleWidtht As Long
    MaxTriangleWidtht = GetMaxItemValue(triangleWidth) / 3
    
    'Left Side
    Dim triangleTopLeftName As String
    triangleTopLeftName = "Customer Management"
    Dim triangleTopLeft As Shape
    
    
    Dim triangleTopLeft1Name As String
    triangleTopLeft1Name = "Sales"
    Dim triangleTopLeft1 As Shape
    
    Dim triangleBottomLeftName As String
    triangleBottomLeftName = "Services"
    Dim triangleBottomLeft As Shape
    
    Dim triangleBottomLeft1Name As String
    triangleBottomLeft1Name = "Products and provisioning"
    Dim triangleBottomLeft1 As Shape

    'Right Side
    Dim triangleTopRightName As String
    triangleTopRightName = "Online"
    Dim triangleTopRight As Shape
    
    Dim triangleTopRight1Name As String
    triangleTopRight1Name = "Bill-IT"
    Dim triangleTopRight1 As Shape
    
    Dim triangleBottomRightName As String
    triangleBottomRightName = "Finance"
    Dim triangleBottomRight As Shape
    
    Dim triangleBottomRight1Name As String
    triangleBottomRight1Name = "Products and Configuration Management"
    Dim triangleBottomRight1 As Shape
    
    Dim leftValue As Long
    leftValue = 50
    Dim topValue As Long
    topValue = 50
    
    'Top Left side
    Set triangleTopLeft = wksht.Shapes.AddShape(msoShapeRightTriangle, leftValue, topValue, MaxTriangleWidtht, MaxTriangleHeight)
    triangleTopLeft.name = triangleTopLeftName
    Set triangleTopLeft1 = wksht.Shapes.AddShape(msoShapeRightTriangle, leftValue, topValue, MaxTriangleWidtht, MaxTriangleHeight)
    triangleTopLeft1.name = triangleTopLeft1Name
    triangleTopLeft1.Flip msoFlipHorizontal
    triangleTopLeft1.Flip msoFlipVertical
    
    'Botton Left side
    Set triangleBottomLeft = wksht.Shapes.AddShape(msoShapeRightTriangle, leftValue, topValue + triangleTopLeft1.height, triangleTopLeft.width, triangleTopLeft.height)
    triangleBottomLeft.name = triangleBottomLeftName
    triangleBottomLeft.Flip msoFlipHorizontal
    Set triangleBottomLeft1 = wksht.Shapes.AddShape(msoShapeRightTriangle, leftValue, topValue + triangleTopLeft1.height, triangleTopLeft.width, triangleTopLeft.height)
    triangleBottomLeft1.name = triangleBottomLeft1Name
    triangleBottomLeft1.Flip msoFlipVertical

    'Top Right side
    Set triangleTopRight = wksht.Shapes.AddShape(msoShapeRightTriangle, leftValue + triangleTopLeft1.width, topValue, MaxTriangleWidtht, MaxTriangleHeight)
    triangleTopRight.name = triangleTopRightName
    triangleTopRight.Flip msoFlipHorizontal
    
    Set triangleTopRight1 = wksht.Shapes.AddShape(msoShapeRightTriangle, leftValue + triangleTopLeft1.width, topValue, MaxTriangleWidtht, MaxTriangleHeight)
    triangleTopRight1.name = triangleTopRight1Name
    triangleTopRight1.Flip msoFlipVertical
  
    'Bottom right side
    Set triangleBottomRight = wksht.Shapes.AddShape(msoShapeRightTriangle, leftValue + triangleTopRight1.width, topValue + triangleTopRight1.height, MaxTriangleWidtht, MaxTriangleHeight)
    triangleBottomRight.name = triangleBottomRightName
    Set triangleBottomRight1 = wksht.Shapes.AddShape(msoShapeRightTriangle, leftValue + triangleTopRight1.width, topValue + triangleTopRight1.height, MaxTriangleWidtht, MaxTriangleHeight)
    triangleBottomRight1.name = triangleBottomRight1Name
    triangleBottomRight1.Flip msoFlipHorizontal
    triangleBottomRight1.Flip msoFlipVertical

    Set triangleTopLeft = Nothing
    Set triangleTopLeft1 = Nothing
    Set triangleBottomLeft = Nothing
    Set triangleBottomLeft1 = Nothing
    Set triangleTopRight = Nothing
    Set triangleTopRight1 = Nothing
    Set triangleBottomRight = Nothing
    Set triangleBottomRight1 = Nothing
	Set triangleHeight = Nothing
	Set triangleWidth  = Nothing
End Sub

Function GetMaxItemValue(myDict As Dictionary) As Variant
    Dim maxItem As Variant
    Dim maxValue As Variant
    Dim item As Variant
    For Each item In myDict.Items
        If maxValue < item Then
            maxValue = item
        End If
    Next
    GetMaxItemValue = maxValue
End Function

Function TeamMapListOfSystemsInAProject()
    Dim dataSheet As Worksheet
    Set dataSheet = Worksheets(DATA_SHEET_NAME)
    
    'Retrieve list of all project from Data sheet even if the cell is empty
    Dim ListOfAllProject As New Dictionary
    Set ListOfAllProject = Sheet3.RetrieveListOfAllProject

    ''''******Test Type DICTIONARY******''''
    Dim TestTypeDict As New Dictionary
    Set TestTypeDict = Sheet3.RetrieveListOfAllTestTypeByProject
    
    ''''******Retrieve list of system
    Dim ListOfAllSystemDict As New Dictionary
    Set ListOfAllSystemDict = Sheet3.RetrieveListOfAllSystem
    
    ''''*************Determine which system are tested by team/testtype
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim projectName As Variant
    Dim ProjectSystemName As Variant
    Dim ListOfProjectBySystemsDict As New Dictionary
    Dim listOfSystemsByProjectArr As New Collection
    'The 3 For-loop below works as expected
    For j = 0 To TestTypeDict.Count - 1
        For i = 0 To ListOfAllProject.Count - 1
            'Make sure that we are under the same column in projectName and test type
            If StrComp(Sheet3.Cell_Position(TestTypeDict.Keys()(j), "COLUMN_ID"), Sheet3.Cell_Position(ListOfAllProject.Keys()(i), "COLUMN_ID"), vbTextCompare) = 0 Then
                projectName = ListOfAllProject.Items()(i) '& " (" & TestTypeDict.Items()(j) & ")"
                If Not ListOfProjectBySystemsDict.Exists(projectName) Then
                    Set listOfSystemsByProjectArr = New Collection ' Clear the collection
                    ListOfProjectBySystemsDict.Add projectName, listOfSystemsByProjectArr
                End If
                
                For k = 0 To ListOfAllSystemDict.Count - 1
                    'Verify that cell with the coordinate projectName/TestType column name  and systemName row nr have a value
                    If Len(dataSheet.Range(Sheet3.Cell_Position(TestTypeDict.Keys()(j), "COLUMN_ID") & Sheet3.Cell_Position(ListOfAllSystemDict.Keys()(k), "ROW_NR")).Value) > 0 Then
                        ProjectSystemName = ListOfAllSystemDict.Items()(k)
                        
                        If listOfSystemsByProjectArr.Count > 0 Then
                            If Not CollectionContains(listOfSystemsByProjectArr, ProjectSystemName) Then
                                listOfSystemsByProjectArr.Add ProjectSystemName
                            End If
                        Else
                            listOfSystemsByProjectArr.Add ProjectSystemName
                        End If
                    End If
                Next k
            End If
        Next i
    Next j
    Set TeamMapListOfSystemsInAProject = ListOfProjectBySystemsDict
End Function
   
Function CollectionContains(coll As Collection, target As Variant) As Boolean
    Dim i As Long
    For i = 1 To coll.Count
        If coll(i) = target Then
            CollectionContains = True
            Exit Function
        End If
    Next i
    CollectionContains = False
End Function
   
Function GetShapeMeasurements() As Variant
    Dim wksht As Worksheet
    Set wksht = Worksheets(SYSTEM_COVERAGE_SHEET_NAME)
    
    Dim dict As New Dictionary
    Dim results() As Variant
    ReDim results(1 To 4, 1 To wksht.Shapes.Count)
    
    Dim i As Long
    For i = 1 To wksht.Shapes.Count
        Dim sh As Shape
        Set sh = wksht.Shapes(i)
        
        results(1, i) = sh.height
        results(2, i) = sh.width
        results(3, i) = sh.Rotation       
        dict(sh.name) = i
    Next i
    
    ' Return the results as a dictionary and an array
    Dim returnVal(0 To 1) As Variant
    returnVal(0) = results
    Set returnVal(1) = dict
    GetShapeMeasurements = returnVal
End Function

Using the function above, create a Main function that position each system rectangle shape to its corresponding triangle project shape
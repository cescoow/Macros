' Based on https://modthemachine.typepad.com/my_weblog/2011/06/writing-work-points-to-an-excel-file.html
' Coordinates based on the part coordinate system
' Modified to write the coordinates in separated colums 
' First colum is the point name, #2 is the X, 3# the Y, 4# the Z 
Public Sub ExportWorkPoints()
    ' Get the active part document.
    Dim partDoc As PartDocument
    If ThisApplication.ActiveDocumentType = kPartDocumentObject Then
        Set partDoc = ThisApplication.ActiveDocument
    Else
        MsgBox "A part must be active."
        Exit Sub
    End If
    
    ' Check to see if any work points are selected.
    Dim points() As WorkPoint
    Dim pointCount As Long
    pointCount = 0
    If partDoc.SelectSet.Count > 0 Then
        ' Dimension the array so it can contain the full
        ' list of selected items.
        ReDim points(partDoc.SelectSet.Count - 1)
        
        Dim selectedObj As Object
        For Each selectedObj In partDoc.SelectSet
            If TypeOf selectedObj Is WorkPoint Then
                Set points(pointCount) = selectedObj
                pointCount = pointCount + 1
            End If
        Next
        
        ReDim Preserve points(pointCount - 1)
    End If
    
    ' Ask to see if it should operate on the selected points
    ' or all points.
    Dim getAllPoints As Boolean
    getAllPoints = True
    If pointCount > 0 Then
        Dim result As VbMsgBoxResult
        result = MsgBox("Some work points are selected.  " & _
                "Do you want to export only the " & _
                "selected work points?  (Answering " & _
                """No"" will export all work points)", _
                vbQuestion + vbYesNoCancel)
        If result = vbCancel Then
            Exit Sub
        End If
    
        If result = vbYes Then
            getAllPoints = False
        End If
    Else
        If MsgBox("No work points are selected.  All work points" & _
                  " will be exported.  Do you want to continue?", _
                  vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    
    Dim partDef As PartComponentDefinition
    Set partDef = partDoc.ComponentDefinition
    If getAllPoints Then
        ReDim points(partDef.WorkPoints.Count - 2)
        
        ' Get all of the workpoints, skipping the first,
        ' which is the origin point.
        Dim i As Integer
        For i = 2 To partDef.WorkPoints.Count
            Set points(i - 2) = partDef.WorkPoints.Item(i)
        Next
    End If
    
    ' Get the filename to write to.
    Dim dialog As FileDialog
    Dim filename As String
    Call ThisApplication.CreateFileDialog(dialog)
    With dialog
        .DialogTitle = "Specify Output .CSV File"
        .Filter = "Comma delimited file (*.csv)|*.csv"
        .FilterIndex = 0
        .OptionsEnabled = False
        .MultiSelectEnabled = False
        .ShowSave
        filename = .filename
    End With
    
    If filename <> "" Then
        ' Write the work point coordinates out to a csv file.
        On Error Resume Next
        Open filename For Output As #1
        If Err.Number <> 0 Then
            MsgBox "Unable to open the specified file. " & _
                   "It may be open by another process."
            Exit Sub
        End If
        
        ' Get a reference to the object to do unit conversions.
        Dim uom As UnitsOfMeasure
        Set uom = partDoc.UnitsOfMeasure
        
        ' Write the points, taking into account the current default
        ' length units of the document.
        For i = 0 To UBound(points)
            Dim xCoord As Double
            xCoord = uom.ConvertUnits(points(i).Point.X, _
                 kCentimeterLengthUnits, kDefaultDisplayLengthUnits)
                     
            Dim yCoord As String
            yCoord = uom.ConvertUnits(points(i).Point.Y, _
                 kCentimeterLengthUnits, kDefaultDisplayLengthUnits)
                     
            Dim zCoord As String
            zCoord = uom.ConvertUnits(points(i).Point.Z, _
                 kCentimeterLengthUnits, kDefaultDisplayLengthUnits)
                     
            Print #1, (i) & ";" & _
                Format(xCoord) & ";" & _
                Format(yCoord) & ";" & _
                Format(zCoord)
        Next
        
        Close #1
        
        MsgBox "Finished writing data to """ & filename & """"
    End If
End Sub


Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
#If VBA7 Then ' Excel 2010 or later
 
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)
 
#Else ' Excel 2007 or earlier
 
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As Long)
 
#End If

Dim CWDIR As String
'insert flange icon
Sub Flange()
    Call PlaceImage("/Flange.png", 11, 3)
End Sub

'insert bell icon
Sub Bell()
    Call PlaceImage("/Bell Reducer.png", 12, 10)
End Sub

'insert sudden icon
Sub Sudden()
    Call PlaceImage("/Sudden Reducer.png", 12, 10)
End Sub

'insert elbow icon
Sub Elbow()
    Call PlaceImage("/Elbow.png", 16, 16)
End Sub

'insert tee icon
Sub Tee()
    Call PlaceImage("/Tee.png", 28, 20)
End Sub

'insert cross icon
Sub Cross()
    Call PlaceImage("/cross.png", 29, 29)
End Sub

'insert straight pipe shape
Sub Straight()
    Call PlaceImage("/Straight.png", 30, 11)
End Sub
'insert 90 degree sweep shape
Sub Sweep()
    Call PlaceImage("/90 Sweep.png", 25, 25)
End Sub
'insert switch shape
Sub Switch()
    Call PlaceImage("/Switch.png", 12, 12)
End Sub

'insert pressure port icon
Sub pressurePort()
    Call PlaceImage("/Pressure Port.png", 6, 12)
End Sub

'insert pressure tank shape
Sub pressureTank()
    Call PlaceImage("/Pressure Tank.png", 20, 30)
End Sub
'insert inline pump shape
Sub Inline()
    Call PlaceImage("/Inline.png", 15, 25)
End Sub
'insert PMA shape
Sub PMA()
    Call PlaceImage("/PMA.png", 25, 45)
    If Worksheets("SuperSecretData").Cells(13, 2) = 0 Then  'select the side the discharge is, based on whether this is for the large or small tank, north or south
         'large tank, do nothing
    ElseIf Worksheets("SuperSecretData").Cells(13, 2) = 1 Then
        ActiveSheet.Shapes(ActiveSheet.Shapes.Count).Flip msoFlipHorizontal 'small tank
    Else
        ActiveSheet.Shapes(ActiveSheet.Shapes.Count).Flip msoFlipHorizontal
        ActiveSheet.Shapes(ActiveSheet.Shapes.Count).Flip msoFlipVertical 'small tank south side
    End If
End Sub
'insert union shape
Sub Union()
    Call PlaceImage("/Union.png", 13, 17)
End Sub
'insert dimension line shape
Sub dimensionLine()
    Call PlaceImage("/Line.png", 1, 28)
End Sub
'insert dimension arrow head shape
Sub dimensionArrowHead()
    Call PlaceImage("/Arrow Head.png", 6, 9)
End Sub

'insert Foot Valve icon
Sub footValve()
    Call PlaceImage("/Foot Valve.png", 8, 23)
End Sub

'insert Ball Valve icon
Sub ballValve()
    Call PlaceImage("/Ball Valve.png", 15, 26)
End Sub

'insert Check Valve icon
Sub checkValve()
    Call PlaceImage("/Check Valve.png", 8, 17)
End Sub

'insert flow switch icon
Sub flowSwitch()
    Call PlaceImage("/Flow Switch.png", 10, 10)
End Sub

'insert full CAM icon
Sub fullCAM()
    Call PlaceImage("/CAM Full.png", 18, 21)
End Sub

'insert female CAM icon
Sub femaleCAM()
    Call PlaceImage("/CAM Female.png", 18, 18)
End Sub

'insert male CAM icon
Sub maleCAM()
    Call PlaceImage("/CAM Male.png", 8, 9)
End Sub

'insert relief valve icon
Sub reliefValve()
    Call PlaceImage("/Relief Valve.png", 27, 12)
End Sub

'insert gate valve shape
Sub gateValve()
    Call PlaceImage("/Gate Valve.png", 13, 20)
End Sub

'insert coupling icon
Sub coupling()
    Call PlaceImage("/Coupling.png", 9, 10)
End Sub

'insert WEG shape
Sub WEG()
    Call PlaceImage("/WEG.png", 48, 162)
End Sub

'insert MH Series shape
Sub MH()
    Call PlaceImage("/MH Series.png", 100, 41)
End Sub

'insert FTB Series shape
Sub FTB()
    Call PlaceImage("/FTB Series.png", 100, 40)
End Sub

'insert FCE Series shape
Sub FCE()
    Call PlaceImage("/FCE Series.png", 77, 45)
End Sub

'insert BT4 Series shape
Sub BT()
    Call PlaceImage("/BT4 Series.png", 120, 50)
End Sub



Sub dimensionLength(Optional errorCount As Long = 0)
    On Error GoTo errHandler
    
    '''''''''''''''''''''''''''''''''''''''''define variables
    Dim ws As Worksheet
    Dim imagePath As String
    Dim imgRotation As Double
    CWDIR = Sheets("SuperSecretData").Range("D1")  'get the Current Working Directory
    imagePath = CWDIR + "/Dimension Block.png"  'the dimension block picture file
    Dim selectedImage As Shape
    Dim dimensionImage As Shape
    Dim dimensionLineL As Shape
    Dim dimensionLineLFull As Shape
    Dim dimensionLineRFull As Shape
    Dim dimensionLineR As Shape
    Dim conLineL As Shape
    Dim conLineR As Shape
    Dim groupShape As Shape
    Dim dimBlock As Shape
    Dim rotateNumber As Integer
    Dim alignOffset As Double
    Dim originalX As Double
    Dim originalY As Double

    Set ws = ActiveSheet
    Set selectedImage = GetShape  'get selected shape and assign it to the variable
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''conditions if an image is selected
    If Not selectedImage Is Nothing Then
        'the number, always horizontal
        imgRotation = selectedImage.Rotation  'get selected image's rotation and assign it to a variable
        Worksheets("SuperSecretData").Range("F23").Columns.AutoFit   'adjusts column width to fit the data
        Sheets("SuperSecretData").Range("F23").CopyPicture   'Copies cell as picture, thus allowing better manueverability of everything
        ws.Range("I16").PasteSpecial  'pastes special as picture
        Set dimensionImage = GetShape  'set it as the pasted picture, which should be selected at this point
        
        If (imgRotation > 90 And imgRotation < 270) Then  'flip it if it is upside-down, that way the dimensioning text is nto upside-down also.
            dimensionImage.Rotation = dimensionImage.Rotation + 180
        End If

        Set dimensionLineL = ws.Shapes.AddLine(selectedImage.Left, selectedImage.Top + selectedImage.height, selectedImage.Left, selectedImage.Top + selectedImage.height + 10)  'create line so the top connects to the dimension line
        Set dimensionLineLFull = ws.Shapes.AddLine(selectedImage.Left, selectedImage.Top + selectedImage.height, selectedImage.Left, selectedImage.Top + selectedImage.height + 20)  'create the full dimension line
        Set dimensionLineR = ws.Shapes.AddLine(selectedImage.Left + selectedImage.width, selectedImage.Top + selectedImage.height, selectedImage.Left + selectedImage.width, selectedImage.Top + selectedImage.height + 10)  'create line so the top connects to the dimension line
        Set dimensionLineRFull = ws.Shapes.AddLine(selectedImage.Left + selectedImage.width, selectedImage.Top + selectedImage.height, selectedImage.Left + selectedImage.width, selectedImage.Top + selectedImage.height + 20)  'create the full dimension line
        dimensionImage.Top = selectedImage.Top + selectedImage.height + 1 'set the dimension image to be in the center
        dimensionImage.Left = selectedImage.Left + selectedImage.width / 2 - dimensionImage.width / 2   'set to be in center
        rotateNumber = 0

        Set conLineR = ws.Shapes.AddConnector(msoConnectorStraight, 1, 1, 1, 1)  'make a straight line connector
        With conLineR:
            .ConnectorFormat.BeginConnect connectedshape:=dimensionImage, connectionsite:=4  'right center of dimension image  '1 = top, 2=left, 3=bottom, 4 = right
            .ConnectorFormat.EndConnect connectedshape:=dimensionLineR, connectionsite:=2  'top of the line
            .Line.ForeColor.RGB = RGB(0, 0, 0)  'set line color to black
            .Line.EndArrowheadStyle = msoArrowheadTriangle  'set the end to be a triangle point
            While .ZOrderPosition > 2
                .ZOrder msoSendBackward  'send the line to the back layer, that way it is not over the part
            Wend
        End With
        Set conLineL = ws.Shapes.AddConnector(msoConnectorStraight, 1, 1, 1, 1)  'make a straight line connector
        With conLineL:
            .ConnectorFormat.BeginConnect connectedshape:=dimensionImage, connectionsite:=2  'left center of dimension image  '1 = top, 2=left, 3=bottom, 4 = right
            .ConnectorFormat.EndConnect connectedshape:=dimensionLineL, connectionsite:=2  'top of the line
            .Line.ForeColor.RGB = RGB(0, 0, 0)  'set line color to black
            .Line.EndArrowheadStyle = msoArrowheadTriangle   'set the end to be a triangle point
            While .ZOrderPosition > 2
                .ZOrder msoSendBackward  'send the line to the back layer, that way it is not over the part
            Wend
        End With
        
        Dim dimLine As Shape
        For Each dimLine In ws.Shapes   'iterate through all the shapes, color the lines black
            If dimLine.Type = msoLine Then dimLine.Line.ForeColor.RGB = RGB(0, 0, 0)
        Next dimLine
        
        originalX = selectedImage.Left  'sets initial X value
        originalY = selectedImage.Top  'sets initial Y value
        Set dimBlock = ws.Shapes.Range(Array(dimensionLineL.Name, dimensionLineLFull.Name, dimensionLineR.Name, dimensionLineRFull.Name, conLineR.Name, conLineL.Name, dimensionImage.Name)).Group
        imgRotation = selectedImage.Rotation  'get selected image's rotation and assign it to a variable
        selectedImage.Rotation = 0  'set image to straight horizontal
        dimBlock.Top = selectedImage.Top + selectedImage.height   'set the dimensioning block's top to the selected image's bottom
        Set groupShape = ws.Shapes.Range(Array(dimBlock.Name, selectedImage.Name)).Group  'group both the dimensioning block and the selected image
        groupShape.Rotation = imgRotation  'rotate both dimension and selected image together
        groupShape.Ungroup  'ungroup the dimensioning block and selected image.  This allows both to be edited independantly
        dimBlock.Left = dimBlock.Left - selectedImage.Left + originalX  'uses the offset created from grouping to return dimensioning block to original X position
        dimBlock.Top = dimBlock.Top - selectedImage.Top + originalY  'uses the offset created from grouping to return dimensioning block to original Y position
        selectedImage.Left = originalX  'returns selected image to original X
        selectedImage.Top = originalY  'returns selected image to original Y
        
        'MsgBox ("angle:" & imgRotation & vbCrLf & rotateNumber & vbCrLf & selectedImage.width & vbCrLf & alignOffset)  'displays information regarding dimension command.  Keep commented unless debugging.
    Else  'no item selected to dimension!
        MsgBox ("There is no item selected to dimension")
    End If
    On Error GoTo 0
Exit Sub
    
errHandler:   'display a message box with an error description

    MsgBox ("Error occured while adding the dimension length" & vbCrLf & "Error Type: " & Err.Description)  'displays the error description upon an error occuring
    Exit Sub  'exit the sub, just in case the error does not resolve itself

End Sub

Sub descriptionImageScript()
    On Error GoTo errHandler
    Dim ws As Worksheet
    Dim dimensionImage As Shape
    Dim selectedImage As Shape
    Dim conLine As Shape
    
    Set ws = ActiveSheet
    Set selectedImage = GetShape  'assigns the selected image to the variable
    
    If Not selectedImage Is Nothing Then
        Worksheets("SuperSecretData").Range("E23").Columns.AutoFit  'make the column width fit to the text that is currently in it
        Sheets("SuperSecretData").Range("E23").CopyPicture   'Copies cell as picture, thus allowing better manueverability of everything
        ws.Range("I16").PasteSpecial  'pastes special as picture
        Set dimensionImage = GetShape  'the pasted image should be selected at this point
        Set conLine = ws.Shapes.AddConnector(msoConnectorStraight, 1, 1, 1, 1)  'make a straight line connector
        With conLine:
            .ConnectorFormat.BeginConnect connectedshape:=selectedImage, connectionsite:=4  'left center of selected image  '1 = top, 2=left, 3=bottom, 4 = right
            .ConnectorFormat.EndConnect connectedshape:=dimensionImage, connectionsite:=2  'left center of textbox
            .Line.ForeColor.RGB = RGB(0, 0, 0)  'set line color to black
            While .ZOrderPosition > 2
                .ZOrder msoSendBackward  'send the line to the back layer, that way it is not over either the part or the description
            Wend
        End With
    Else 'there was not a selected image to describe
        MsgBox ("There is no item selected to describe")
    End If
    
Exit Sub
    
errHandler:   'display a message box with an error description
MsgBox ("Error occured while adding a description" & vbCrLf & "Error Type: " & Err.Description)

End Sub

Sub commentImageScript()
    On Error GoTo errHandler
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Worksheets("SuperSecretData").Range("E25").Columns.AutoFit  'adjust column to shrink or grow to size with the text
    If Worksheets("SuperSecretData").Range("E25").ColumnWidth > 30 Then 'force the cell to use wrap around if it is longer than 30 units
        Worksheets("SuperSecretData").Range("E25").ColumnWidth = 30
    End If
    Worksheets("SuperSecretData").Range("E25").Rows.AutoFit  'make row autofit to content, so that when it is copied the image is the correct size
    Sheets("SuperSecretData").Range("E25").CopyPicture   'Copies cell as picture, thus allowing better manueverability of everything
    ActiveCell.PasteSpecial  'pastes special as picture
    
Exit Sub
    
errHandler:  'display a message box with an error description
MsgBox ("Error occured while adding a comment" & vbCrLf & "Error Type: " & Err.Description)

End Sub

Sub exportSheet(Optional errorCount As Long = 0)
    On Error Resume Next  'resume the next line upon error
    
    'define variables
    Dim ws As Worksheet 'define variable to hold worksheet
    Set ws = ActiveSheet 'set to current worksheet
    Dim picSnap As Shape
    Dim pic As Variant
    
    If Not sheetExists("Plumbing_Diagram") Then Sheets.Add.Name = "Plumbing_Diagram"   'checks if export sheet exists, creates it if not
    Sheets("Plumbing_Diagram").Move after:=ws   'move the plumbing diagram sheet after the main sheet
    
    ''''''''''''''''''''''''''''''''''''''header
    ws.Range("A1", "N3").Copy  'copy header
    Sheets("Plumbing_Diagram").Range("A1", "N3").PasteSpecial  'maintain formatting
    Sheets("Plumbing_Diagram").Range("A1", "J2").Clear  'clear contents of the cells containing the flow tube and manifold selection and drop down boxes
    Sheets("Plumbing_Diagram").Range("A1", "J2").Interior.Color = vbWhite  'make background color white

    ''''''''''''''''''''''''''''''''''''''diagram
    ws.Range("A4", "N32").CopyPicture _
        Appearance:=xlScreen, _
        Format:=xlBitmap   'Copy the main diagram range.
    If (Sheets("Plumbing_Diagram").Shapes.Count > 0) Then
        Sheets("Plumbing_Diagram").Shapes(Sheets("Plumbing_Diagram").Shapes.Count).Delete 'deletes most recent shape on the export sheet
    End If
    Sheets("Plumbing_Diagram").Pictures.Paste  'paste as picture in export sheet
    Set picSnap = Sheets("Plumbing_Diagram").Shapes(Sheets("Plumbing_Diagram").Shapes.Count)   'assign picSnap to the most recently added shape
    picSnap.Top = Sheets("Plumbing_Diagram").Range("A4").Top  'set where the top of the picture is
    picSnap.Left = Sheets("Plumbing_Diagram").Range("A4").Left  'set where the left of the picture is
    ActiveSheet.Cells(1, 1).Activate 'go back to the upper left cell, clear out the selection range
    Application.CutCopyMode = False  'basically deselect what was just copied
End Sub

Sub exportPicture(Optional errorCount As Long = 0)
    On Error GoTo errHandler
    'define variables
    Dim ws As Worksheet 'define variable to hold worksheet
    Set ws = ActiveSheet 'set to current worksheet
    Dim tmpChart As Chart, picSnap As Shape
    Dim fileSaveName As Variant, pic As Variant

    'Create temporary chart as canvas
    
    ws.Range("A4", "N32").Copy  'Copy the main diagram range.
    ws.Pictures.Paste.Select
    Set picSnap = ws.Shapes(ws.Shapes.Count)   'assign picSnap to the most recently added shape
    Set tmpChart = Charts.Add  'assign new chart to tmpChart
    tmpChart.ChartArea.Clear  'clear all the default chart items out
    tmpChart.Name = "PicChart"  'define the chart's name
    Set tmpChart = tmpChart.Location(Where:=xlLocationAsObject, Name:=ws.Name)  'move chart to main worksheet
    tmpChart.Parent.Top = Range("A4").Top  'set the chart's location to be on top of the material being copied
    tmpChart.Parent.Left = Range("A4").Left
    tmpChart.ChartArea.width = picSnap.width 'set the same width and height of the range selected
    tmpChart.ChartArea.height = picSnap.height
    tmpChart.Parent.Border.LineStyle = 0  'remove chart borders, so it does not interfere with the original image
    
    'Paste range as image to chart
    picSnap.Copy 'copies image
    tmpChart.ChartArea.Select 'select chart area in perperation for pasting the picture
    tmpChart.Paste 'paste the picture in the chart
    
    'Save chart image to file
    fileSaveName = Application.GetSaveAsFilename(InitialFileName:="Plumbing Diagram", fileFilter:="Image (*.jpg), *.jpg")
    tmpChart.Export fileName:=fileSaveName, FilterName:="jpg"
    
    'Clean up
    ws.Cells(1, 1).Activate 'go back to the upper left cell, clear out the selection range
    ws.ChartObjects.Delete 'delete temp chart
    picSnap.Delete 'delete picture of range
    On Error GoTo 0  'return to the start of the subroutine
Exit Sub
    
errHandler:  'display a message box with an error description
    'MsgBox ("Error occured while exporting picture" & vbCrLf & "Error Type: " & Err.Description)
    errorCount = errorCount + 1 'increment the error count
    If errorCount < 5 Then  'less than 5 errors?
        Call exportPicture(errorCount) 'there is an error due to the .Select code sometimes, due to VBA not being an ideal language.  If you try again, it will work
    Else
        Exit Sub  'exit the sub, just in case the error does not resolve itself
    End If
End Sub

'make a hose connector between two selected shapes
Sub hoseConnector()
    On Error GoTo errHandler
    Dim shpRange As ShapeRange
    Dim usrSelection As Variant
    Dim conLine As Shape
    
    Set usrSelection = ActiveWindow.Selection  'get all things selected on the active window
    Set shpRange = usrSelection.ShapeRange   'get the range of shapes selected
    If Not shpRange(2) Is Nothing Then
        Set conLine = ActiveSheet.Shapes.AddConnector(msoConnectorCurve, 1, 1, 1, 1)  'create connector line
        With conLine:
            .ConnectorFormat.BeginConnect connectedshape:=shpRange(1), connectionsite:=3   'set connection site to the original bottom of the picture
            .ConnectorFormat.EndConnect connectedshape:=shpRange(2), connectionsite:=3  'set connection site to the original bottom of the picture
            .Line.ForeColor.RGB = RGB(0, 0, 255)  'set color to blue
            .Line.Weight = 7   'set thickness to 7
            While .ZOrderPosition > 2
                .ZOrder msoSendBackward  'send the line to the back layer, that way it is not over either the part or the description
            Wend
        End With
    Else
        MsgBox ("Two shapes need to be selected in order to connect the hose")
    End If
Exit Sub
    
errHandler:       'display a message box with an error description
MsgBox ("Two shapes need to be selected in order to connect the hose")

End Sub


Sub clearSheet()
    On Error Resume Next
    ActiveSheet.Pictures.Delete  'delete all pictures
    Dim s As Shape  'define shape variable
    For Each s In ActiveSheet.Shapes   'iterate through all the shapes, delete on the line connectors
        If s.Name Like "*Connector*" Then s.Delete
        If s.Name Like "*Group*" Then s.Delete
    Next s
End Sub

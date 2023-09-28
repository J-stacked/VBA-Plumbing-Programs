Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1

Function GetShape() As Shape  'returns a shape based on the current selected image
    If TypeName(Selection) = "Picture" Then Set GetShape = Application.Selection.ShapeRange.Item(1)
End Function

Function sheetExists(sSheet As String) As Boolean  'returns boolean as to whether sheet exists or not
    On Error Resume Next
    sheetExists = (ActiveWorkbook.Sheets(sSheet).Index > 0)
End Function

Sub PlaceImage(fileName As String, width As Integer, height As Integer)
    On Error GoTo errHandler
    Dim ws As Worksheet
    Dim imagePath As String
    Dim imgLeft As Double
    Dim imgTop As Double
    Dim imgScale As Double
    
    CWDIR = Sheets("SuperSecretData").Range("D1")  'get Current Working Directory
    Set ws = ActiveSheet
    imagePath = CWDIR + fileName   'picture file
    'set position to the current active cell
    imgLeft = ActiveCell.left
    imgTop = ActiveCell.top
    imgScale = ws.Range("I1").Value
    
    'Width & Height = -1 means keep original size, dimensions in pixels
    ws.Shapes.AddPicture _
        fileName:=imagePath, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, _
        left:=imgLeft, _
        top:=imgTop, _
        width:=width * imgScale, _
        height:=height * imgScale
        
Exit Sub

errHandler:   'display a message box with an error description
If Err.Number = 1004 Then
    MsgBox ("Error occured while adding the image.  Please ensure OneDrive has backed up the folder before proceeding." & vbCrLf & vbCrLf & "Error Type: " & Err.Description & vbCrLf & "Error Code: " & Err.Number)
Else
    MsgBox ("Error occured while adding the image" & vbCrLf & "Error Type: " & Err.Description & vbCrLf & "Error Code: " & Err.Number)
End If

End Sub



Sub PlaceDot(fileName As String, x As Integer, y As Integer)
    On Error GoTo errHandler
    Dim ws As Worksheet
    Dim imagePath As String
    Dim imgLeft As Double
    Dim imgTop As Double
    
    CWDIR = Sheets("SuperSecretData").Range("D1")  'get Current Working Directory
    Set ws = ActiveSheet
    imagePath = CWDIR + fileName   'picture file
    'set position to the current active cell
    imgLeft = x
    imgTop = y
    
    'Width & Height = -1 means keep original size, dimensions in pixels
    ws.Shapes.AddPicture _
        fileName:=imagePath, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, _
        left:=imgLeft, _
        top:=imgTop, _
        width:=1, _
        height:=1
     
Exit Sub
    
errHandler:   'display a message box with an error description
If Err.Number = 1004 Then
    MsgBox ("Error occured while adding the image.  Please ensure OneDrive has backed up the folder before proceeding." & vbCrLf & vbCrLf & "Error Type: " & Err.Description & vbCrLf & "Error Code: " & Err.Number)
Else
    MsgBox ("Error occured while adding the image" & vbCrLf & "Error Type: " & Err.Description & vbCrLf & "Error Code: " & Err.Number)
End If

End Sub

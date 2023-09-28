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
    imgLeft = ActiveCell.Left
    imgTop = ActiveCell.Top
    imgScale = ws.Range("I1").Value
    
    'Width & Height = -1 means keep original size, dimensions in pixels
    ws.Shapes.AddPicture _
        fileName:=imagePath, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, _
        Left:=imgLeft, _
        Top:=imgTop, _
        width:=width * imgScale, _
        height:=height * imgScale
        
Exit Sub
    
errHandler:   'display a message box with an error description
MsgBox ("Error occured while adding the image" & vbCrLf & "Error Type: " & Err.Description)

End Sub

Sub deleteDimenTry()
    On Error GoTo errHandler
    
    Dim s As Shape  'define shape variable
    For Each s In activehseet.Shapes   'iterate through all the shapes, delete on the line connectors
        If s. Then s.Delete
    Next s
Exit Sub
    
errHandler:   'display a message box with an error description
MsgBox ("Error occured while deleting dimension try" & vbCrLf & "Error Type: " & Err.Description)

End Sub

Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
'insert flange icon
Sub Flange()
    Call PlaceImage("/Flange.png", 11, 3)
End Sub

'insert bell icon
Sub Bell()
    Call PlaceImage("/Bell Reducer.png", 11, 10)
End Sub

'insert solenoid icon
Sub solenoid()
    Call PlaceImage("/Solenoid.png", 41, 29)
End Sub

'insert sudden icon
Sub Sudden()
    Call PlaceImage("/Sudden Reducer.png", 11, 10)
End Sub

'insert 45 icon
Sub FourtyFive()
    Call PlaceImage("/45.png", 10, 10)
End Sub

'insert elbow icon
Sub Elbow()
    Call PlaceImage("/Elbow.png", 11, 11)
End Sub

'insert tee icon
Sub Tee()
    Call PlaceImage("/Tee.png", 15, 11)
End Sub

'insert cross icon
Sub Cross()
    Call PlaceImage("/cross.png", 15, 15)
End Sub

Sub dimensionLine()
    Call PlaceImage("/Line.png", 1, 30)
End Sub

'insert straight pipe shape
Sub Straight()
    Call PlaceImage("/Straight.png", 18, 7)
End Sub

'insert 90 degree sweep shape
Sub Sweep()
    Call PlaceImage("/90 Sweep.png", 20, 20)
End Sub

'insert switch shape
Sub Switch()
    Call PlaceImage("/Switch.png", 12, 12)
End Sub

'insert pressure port icon
Sub pressurePort()
    Call PlaceImage("/Pressure Port.png", 6, 12)
End Sub

'insert pressure gauge icon
Sub pressureGauge()
    Call PlaceImage("/Pressure Gauge.png", 20, 34)
End Sub

'insert funnel icon
Sub funnel()
    Call PlaceImage("/Funnel.png", 32, 30)
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

'insert coupling icon
Sub coupling()
    Call PlaceImage("/Coupling.png", 9, 10)
End Sub

'insert pressure tank shape
Sub pressureTank()
    Call PlaceImage("/Pressure Tank.png", 52, 90)
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

'insert Little Giant WS shape
Sub LittleGiant()
    Call PlaceImage("/Little Giant WS.png", 40, 62)
End Sub

'insert Submersible shape
Sub Submersible()
    Call PlaceImage("/Submersible.png", 20, 156)
End Sub

'insert union shape
Sub Union()
    Call PlaceImage("/Union.png", 13, 17)
End Sub

'insert gate valve shape
Sub gateValve()
    Call PlaceImage("/Gate Valve.png", 13, 20)
End Sub

'insert RMA Top shape
Sub rmaTop()
    Call PlaceImage("/RMA Top.png", 177, 132)
End Sub

'insert RMA Full shape
Sub rmaFull()
    Call PlaceImage("/RMA Full.png", 200, 250)
End Sub

'insert Flow Meter shape
Sub flowMeter()
    Call PlaceImage("/Flow Meter.png", 152, 80)
End Sub

'insert discharge PMA shape
Sub pmaDischarge()
    Call PlaceImage("/PMA Discharge.png", 70, 23)
End Sub

'insert inlet PMA shape
Sub pmaInlet()
    Call PlaceImage("/PMA Inlet.png", 25, 25)
End Sub

'insert SDC shape
Sub SDC()
    Sheets("SDC").Range("A1:N29").Copy (ActiveSheet.Range("A4"))
End Sub

'insert Utility Drive shape
Sub utilityDrive()
    Sheets("Utility Drive").Range("A1:N29").Copy (ActiveSheet.Range("A4"))
End Sub

'insert dimension arrow head shape
Sub dimensionArrowHead()
    Call PlaceImage("/Arrow Head.png", 6, 9)
End Sub

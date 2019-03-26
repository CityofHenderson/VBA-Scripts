'Version 1.0.0.2
'Created by Ian Harshbarger
'for the City of Henderson
'February 28 2019
'GNU General Public License v3

Public Function MANNING(PipeSize As Double, Slope As Double, K As Double, Depth As Double, Coefficient As Double)
Depth = Depth / 12

Dim Radius As Double
Dim Theta As Double
Dim Area As Double
Dim WettedPerimeter As Double
Dim HydraulicRadius As Double
Radius = (PipeSize / 2) / 12

If Depth <= Radius Then
    Theta = 2 * Application.WorksheetFunction.Acos((Radius - Depth) / Radius)
    Area = Radius ^ 2 * (Theta - Sin(Theta)) / 2
    WettedPerimeter = Radius * Theta
    HydraulicRadius = Area / WettedPerimeter
Else
    Depth = 2 * Radius - Depth
    Theta = 2 * Application.WorksheetFunction.Acos((Radius - Depth) / Radius)
    Area = (Application.WorksheetFunction.Pi() * Radius ^ 2) - (Radius ^ 2 * (Theta - Sin(Theta)) / 2)
    WettedPerimeter = (2 * Application.WorksheetFunction.Pi() * Radius) - (Radius * Theta)
    HydraulicRadius = Area / WettedPerimeter
End If
MANNING = ((K * Area * (HydraulicRadius ^ (2 / 3)) * (Slope ^ (1 / 2))) / Coefficient)
'MsgBox "K =" + CStr(K) + " Area = " + CStr(Area) + " HydraulicRadius = " + CStr(HydraulicRadius) + " Slope = " + CStr(Slope) + " Coefficient = " + CStr(Coefficient)

End Function

Sub DescribeFunction()
   Dim FuncName As String         'name of the function you want to register
   Dim FuncDesc As String         'description of the function itself
   Dim Category As String         'description of function arguments
   Dim ArgDesc(1 To 5) As String  'description of function arguments

   FuncName = "MANNING"
   FuncDesc = "Solves for the Manning Formula for a Circlular Pipe"
   Category = 7 'Text category
    ArgDesc(1) = "Nominal Pipe Size"
    ArgDesc(2) = "Slope of the Upstream Sewer Line"
    ArgDesc(3) = "CFS use 1.49, GPM use 669, and MGD use 0.963"
    ArgDesc(4) = "Depth of the liquid"
    ArgDesc(5) = "Typical is 0.013"
   Application.MacroOptions _
      Macro:=FuncName, _
      Description:=FuncDesc, _
      Category:=Category, _
      ArgumentDescriptions:=ArgDesc
End Sub



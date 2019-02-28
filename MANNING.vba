'Version 1.0.0.0
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
    Area = (Application.WorksheetFunction.pi() * Radius ^ 2) - (Radius ^ 2 * (Theta - Sin(Theta)) / 2)
    WettedPerimeter = (2 * Application.WorksheetFunction.pi() * Radius) - (Radius * Theta)
    HydraulicRadius = Area / WettedPerimeter
End If
MANNING = ((K * Area * (HydraulicRadius ^ (2 / 3)) * (Slope ^ (1 / 2))) / Coefficient)
'MsgBox "K =" + CStr(K) + " Area = " + CStr(Area) + " HydraulicRadius = " + CStr(HydraulicRadius) + " Slope = " + CStr(Slope) + " Coefficient = " + CStr(Coefficient)

End Function


Sub RegisterUDF()
Dim strFunc As String   'name of the function you want to register
Dim strDesc As String   'description of the function itself
Dim strArgs() As String 'description of function arguments

    'Register Linterp linear interpolation function
    ReDim strArgs(1 To 5) 'The upper bound is the number of arguments in your function
    strFunc = "Manning"
    strDesc = "Solves for the Manning Formula for a Circlular Pipe"
    strArgs(1) = "Nominal Pipe Size"
    strArgs(2) = "Slope of the Upstream Sewer Line"
    strArgs(3) = "CFS use 1.49, GPM use 669, and MGD use 0.963"
    strArgs(4) = "Depth of the liquid"
    strArgs(5) = "Typical is is 0.013"
    Application.MacroOptions Macro:=strFunc, _
                             Description:=strDesc, _
                             ArgumentDescriptions:=strArgs, _
                             Category:="My Custom Category"
End Sub

Sub UnregisterUDF()
    Application.MacroOptions Macro:="IFERROR", Description:=Empty, Category:=Empty
End Sub

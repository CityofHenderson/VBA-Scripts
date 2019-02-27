Public Function MANNING(PipeSize As Double, Slope As Double, K As Double, Depth As Double, Coefficient As Double)

'PipeSize = 14.426
'Slope = 0.0072
'K = 1.49
'Depth = 8.6556
'Coefficient = 0.013

Dim Radius As Double
Radius = (PipeSize / 2) / 12

Dim Theta As Double
Theta = 2 * Application.WorksheetFunction.Acos((Radius - Depth) / (PipeSize / 2))

Dim Area As Double
Area = Radius ^ 2 * (Theta - Sin(Theta)) / 2
MsgBox Area

Dim WettedPerimeter As Double
WettedPerimeter = Radius * Theta

Dim HydraulicRadius As Double
HydraulicRadius = Area / WettedPerimeter

MANNING = ((K * Area * (Radius ^ (2 / 3)) * (Slope ^ (1 / 2))) / Coefficient)


End Function


Sub RegisterUDF()
Dim strFunc As String   'name of the function you want to register
Dim strDesc As String   'description of the function itself
Dim strArgs() As String 'description of function arguments

    'Register Linterp linear interpolation function
    ReDim strArgs(1 To 5) 'The upper bound is the number of arguments in your function
    strFunc = "Manning"
    strDesc = "2D Linear Interpolation function that automatically picks which range " & _
              "to interpolate between based on the closest KnownX value to the NewX " & _
              "value you want to interpolate for."
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






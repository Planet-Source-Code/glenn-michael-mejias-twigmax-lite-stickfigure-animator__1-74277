Attribute VB_Name = "Module1"
Public Const PI = 3.1415
Public Type POINTAPI
    x   As Double
    y   As Double
End Type
Public Function rotatepoint(angle As Double, LLenght As Double, x As Double, y As Double) As POINTAPI
    Dim Radian As Double
    sPIDiv = 0.017452
    Radian = angle * sPIDiv
    rotatepoint.x = (Cos(Radian) * LLenght) + x
    rotatepoint.y = (Sin(Radian) * LLenght) + y
End Function
Public Function getangle(x1 As Double, y1 As Double, x2 As Double, y2 As Double) As Double
    Dim theta As Double
    Dim m1 As Double
    Dim m2 As Double
    m1 = x2 - x1
    m2 = y2 - y1
    If m1 = 0 And m2 > 0 Then
        getangle = 90
    ElseIf m1 = 0 And m2 < 0 Then
        getangle = 270
    ElseIf m1 > 0 Then
        theta = m2 / m1
        getangle = Atn(theta) / PI * 180
    ElseIf m1 < 0 Then
        theta = m2 / m1
        getangle = (Atn(theta) / PI * 180) - 180
    End If
    If getangle < 0 Then
        getangle = getangle + 360
    End If
End Function

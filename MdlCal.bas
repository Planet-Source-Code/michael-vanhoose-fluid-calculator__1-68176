Attribute VB_Name = "MdlCal"
Public ReynoldsNum As Double
Public bturb As Boolean
Public f As Double
Public Function Velocitys(GPM As Double, d As Double) As Double
    Velocitys = GPM * 0.4085 / (d ^ 2)
    
    If Velocitys < 5 Then
        FrmMain.LblComments.Caption = "Warning:" & vbCr & "Velocitys low value"
    ElseIf Velocitys > 12 Then
        FrmMain.LblComments.Caption = "Warning:" & vbCr & "Velocitys High value"
    End If
End Function
Public Function GetFrictionFromKFactor(V As Double, K As Double) As Double
    GetFrictionFromKFactor = K * ((V ^ 2) / (2 * 32.174))
End Function
Public Function GetHorsepowertoRaiseWater(GPM As Double, SG As Double, FT As Double) As Double
    GetHorsepowertoRaiseWater = (GPM * SG * FT) / 3960
End Function
Public Function GetWeightofWater(L As Double, d As Double) As Double
    GetWeightofWater = L * (d ^ 2) * 0.34
End Function
Public Function GetCapacity(A As Double, V As Double) As Double
    GetCapacity = 448.83 * (A * V)
End Function
Public Function Getreducer(d1large As Double, d2small As Double) As Double
    Getreducer = (0.8 * (1 - ((d2small / d1large) ^ 2)) * Sin((45 / 2) * 3.14159265358979 / 180)) / ((d2small / d1large) ^ 4)
End Function
Public Function Pipesize(GPM As Double, V As Double) As Double
    Pipesize = GPM * 0.002228 / V
End Function
Public Function PipeDia(Area As Double) As Double
    PipeDia = Sqr((Area * (4 / 3.141592654)))
End Function
Public Function Reynolds(Q As Double, p As Double, d As Double, u As Double) As Double
Dim Temp As Double
Temp = 50.6 * ((Q * p) / (d * u))

Select Case Temp
Case Is > 4000
    FrmMain.LblComments.Caption = "Turbulent Flow"
    bturb = True
Case Is < 2000
    FrmMain.LblComments.Caption = "Laminar Flow"
    bturb = False
Case Else
    FrmMain.LblComments.Caption = "Warning:" & vbCr & "Close to Turbulent and Laminar Flow"
    bturb = False
End Select

ReynoldsNum = Temp
Reynolds = Temp
End Function
Public Function Reynoldsold(Velocity As Double, d As Double, V As Double) As Double
Dim Temp As Double
Temp = Velocity * (d / 12) / V

Select Case Temp
Case Is > 4000
    FrmMain.LblComments.Caption = "Turbulent Flow"
    bturb = True
Case Is < 2000
    FrmMain.LblComments.Caption = "Laminar Flow"
    bturb = False
Case Else
    FrmMain.LblComments.Caption = "Warning:" & vbCr & "Close to Turbulent and Laminar Flow"
    bturb = False
End Select

ReynoldsNum = Temp
Reynoldsold = Temp
End Function
Public Function ReynoldsCheck() As Double
    ReynoldsCheck = 64 / ReynoldsNum
    
End Function
Static Function Log10(X)
   Log10 = Log(X) / Log(10#)
End Function
Public Function Colebrook(e As Double, d As Double) As Double
Dim rynolds As Double
Dim check As Double
Dim checkWith As Double

f = ReynoldsCheck()
f = Round(f, 4)
If f = 0 Then f = 0.0001

If bturb = False Then
    checkWith = Round(-2 * (Log10((e / (3.7 * (d / 12))) + (2.51 / (ReynoldsNum * Sqr(f))))), 4)
        Colebrook = checkWith
            Exit Function
End If


checkWith = Round(-2 * (Log10((e / (3.7 * (d / 12))) + (2.51 / (ReynoldsNum * Sqr(f))))), 4)
check = 1 / (Sqr(f))

Do While check > checkWith


    
    checkWith = Round(-2 * (Log10((e / (3.7 * (d / 12))) + (2.51 / (ReynoldsNum * Sqr(f))))), 4)
    check = 1 / (Sqr(f))
    'check = Round(check, 3)

    f = f + 0.0001

'Debug.Print check
'Debug.Print checkWith

Loop

Colebrook = f
End Function
Public Function Darcy(L As Double, d As Double, V As Double) As Double
    Darcy = f * (L / (d / 12)) * (V ^ 2 / (2 * 32.174))
End Function
Public Function GetContractionBig() As Double
Dim d1 As String
Dim d2 As String
d1 = InputBox("Small Diameter in inches")
d2 = InputBox("Large Diameter in inches")
ang = InputBox("Angle of contraction")
    GetContractionBig = ((0.8 * Sin(3.14 / 180 * ang / 2)) * (1 - ((d1 / d2) ^ 2))) / ((d1 / d2) ^ 4)
        MsgBox "K Factor is = " & GetContractionBig
End Function
Public Function GetEnlargementBig() As Double
Dim d1 As String
Dim d2 As String
d1 = InputBox("Small Diameter in inches")
d2 = InputBox("Large Diameter in inches")
ang = InputBox("Angle of contraction")
    GetEnlargementBig = ((2.6 * Sin(3.14 / 180 * ang / 2)) * ((1 - ((d1 / d2) ^ 2)) ^ 2)) / ((d1 / d2) ^ 4)
        MsgBox "K Factor is = " & GetEnlargementBig
End Function
Public Function GetCvFromK(d As Double, K As Double) As Double
    GetCvFromK = (29.9 * (d ^ 2)) / (Sqr(K))
End Function
Public Function GetKFromCv(d As Double, Cv As Double) As Double
    GetKFromCv = (894 * (d ^ 4)) / (Cv ^ 2)
End Function


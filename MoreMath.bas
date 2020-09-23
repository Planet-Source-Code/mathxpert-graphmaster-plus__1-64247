Attribute VB_Name = "MoreMath"
Attribute VB_Description = "Procedures used to perform more mathematical and scientific operations"
' Visual Basic uses radians for the trigonometry
Option Explicit

' Modulo (can also handle decimals AND bigger numbers unlike VB's "Mod")
Function Modulo(ByVal Number As Double, ByVal Divisor As Double) As Double
    Modulo = Number - Divisor * Fix(Number / Divisor)
End Function

' Pi
Function Pi() As Double
Attribute Pi.VB_Description = "Returns pi"
   Pi = Atn(1) * 4
End Function

' Sine in Degrees
Function SinD(ByVal Number As Double) As Double
Attribute SinD.VB_Description = "Returns the sine of a number in degrees"
   SinD = Sin(Deg2Rad(Number))
End Function

' Cosine in Degrees
Function CosD(ByVal Number As Double) As Double
Attribute CosD.VB_Description = "Returns the cosine of a number in degrees"
   CosD = Cos(Deg2Rad(Number))
End Function

' Tangent in Degrees
Function TanD(ByVal Number As Double) As Double
Attribute TanD.VB_Description = "Returns the tangent of a number in degrees"
   TanD = Tan(Deg2Rad(Number))
End Function

' Sine in Gradians
Function SinG(ByVal Number As Double) As Double
Attribute SinG.VB_Description = "Returns the sine of a number in gradians"
   SinG = Sin(Grad2Rad(Number))
End Function

' Cosine in Gradians
Function CosG(ByVal Number As Double) As Double
Attribute CosG.VB_Description = "Returns the cosine of a number in gradians"
   CosG = Cos(Grad2Rad(Number))
End Function

' Tangent in Gradians
Function TanG(ByVal Number As Double) As Double
Attribute TanG.VB_Description = "Returns the tangent of a number in gradians"
   TanG = Tan(Grad2Rad(Number))
End Function

' Inverse Sine
Function Asn(ByVal Number As Double) As Double
Attribute Asn.VB_Description = "Returns the arcsine of a number"
   If Number = 1 Then
      Asn = Pi / 2
   ElseIf Number = -1 Then
      Asn = -Pi / 2
   Else
      Asn = Atn(Number / Sqr(-Number * Number + 1))
   End If
End Function

' Inverse Cosine
Function Acs(ByVal Number As Double) As Double
Attribute Acs.VB_Description = "Returns the arccosine of a number"
   If Number = 1 Then
      Acs = 0
   ElseIf Number = -1 Then
      Acs = Pi
   Else
      Acs = Atn(-Number / Sqr(-Number * Number + 1)) + Pi / 2
   End If
End Function

'We already have the inverse tangent function. It's built into Visual Basic.
'It's called the Atn() function.


' Inverse Sine in Degrees
Function AsnD(ByVal Number As Double) As Double
Attribute AsnD.VB_Description = "Returns the arcsine of a number in degrees"
   AsnD = Rad2Deg(Asn(Number))
End Function

' Inverse Cosine in Degrees
Function AcsD(ByVal Number As Double) As Double
Attribute AcsD.VB_Description = "Returns the arccosine of a number in degrees"
   AcsD = Rad2Deg(Acs(Number))
End Function

' Inverse Tangent in Degrees
Function AtnD(ByVal Number As Double) As Double
Attribute AtnD.VB_Description = "Returns the arctangent of a number in degrees"
   AtnD = Rad2Deg(Atn(Number))
End Function

' Inverse Sine in Gradians
Function AsnG(ByVal Number As Double) As Double
Attribute AsnG.VB_Description = "Returns the arcsine of a number in gradians"
   AsnG = Rad2Grad(Asn(Number))
End Function

' Inverse Cosine in Gradians
Function AcsG(ByVal Number As Double) As Double
Attribute AcsG.VB_Description = "Returns the arccosine of a number in gradians"
   AcsG = Rad2Grad(Acs(Number))
End Function

' Inverse Tangent in Gradians
Function AtnG(ByVal Number As Double) As Double
Attribute AtnG.VB_Description = "Returns the arctangent of a number in gradians"
   AtnG = Rad2Grad(Atn(Number))
End Function

' Base-10 Logarithm
Function Log10(ByVal Number As Double) As Double
Attribute Log10.VB_Description = "Returns the base-10 logarithm of a number"
   Log10 = Log(Number) / Log(10)
End Function

' Reciprocal
Function Recip(ByVal Number As Double) As Double
Attribute Recip.VB_Description = "Returns the reciprocal of a number"
   Recip = 1 / Number
End Function

' N-root
Function NRoot(ByVal Number As Double, ByVal Exponent As Double) As Double
Attribute NRoot.VB_Description = "Returns the base number raised to a specified power of a number"
   NRoot = Number ^ (1 / Exponent)
End Function

' Exponent
Function Exponent(ByVal Number As Double, ByVal Base As Double)
Attribute Exponent.VB_Description = "Returns the exponent raised from a specified base number of a number"
   Exponent = Log(Number) / Log(Base)
End Function

' Degrees to Radians
Function Deg2Rad(ByVal Number As Double) As Double
Attribute Deg2Rad.VB_Description = "Returns the radians of a number in degrees"
   Deg2Rad = Number * Pi / 180
End Function

' Radians to Degrees
Function Rad2Deg(ByVal Number As Double) As Double
Attribute Rad2Deg.VB_Description = "Returns the degrees of a number in radians"
   Rad2Deg = Number * 180 / Pi
End Function

' Degrees to Gradians
Function Deg2Grad(ByVal Number As Double) As Double
Attribute Deg2Grad.VB_Description = "Returns the gradians of a number in degrees"
   Deg2Grad = Number * 10 / 9
End Function

' Gradians to Degrees
Function Grad2Deg(ByVal Number As Double) As Double
Attribute Grad2Deg.VB_Description = "Returns the degrees of a number in gradians"
   Grad2Deg = Number * 9 / 10
End Function

' Radians to Gradians
Function Rad2Grad(ByVal Number As Double) As Double
Attribute Rad2Grad.VB_Description = "Returns the gradians of a number in radians"
   Rad2Grad = Deg2Grad(Rad2Deg(Number))
End Function

' Gradians to Radians
Function Grad2Rad(ByVal Number As Double) As Double
Attribute Grad2Rad.VB_Description = "Returns the radians of a number in gradians"
   Grad2Rad = Deg2Rad(Grad2Deg(Number))
End Function

' Factorial (can also handle decimals like gamma)
Function Factorial(ByVal X As Double) As Double
    Dim i As Long
    Dim f As Double, fi As Double, fo As String, fn As String
    Dim t As Double, dt As Double
    
    Const MSK = "#0.0#############"
    
    If X = Fix(X) Then
        If X < 0 Then Err.Raise 5
        f = 1
        For i = 1 To X: f = f * CDbl(i): Next
    Else
        fi = 170 + (X - Int(X))
        f = 0
        t = 0
        dt = 10 / 3
        t = dt / 2
        
        Do
            fo = Format$(f, MSK)
            f = f + Exp((fi - 1) * Log(t) - t) * dt
            t = t + dt
            fn = Format$(f, MSK)
        Loop Until fo = fn
        
        Do While fi > X + 1.5
            fi = fi - 1
            f = f / fi
        Loop
    End If
    
    Factorial = f
End Function

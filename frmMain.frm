VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GraphMaster Plus"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   9720
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMod 
      Caption         =   "Mod"
      Height          =   375
      Left            =   8880
      TabIndex        =   28
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton cmdAbs 
      Caption         =   "Abs"
      Height          =   375
      Left            =   8040
      TabIndex        =   27
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton cmdGraphSettings 
      Caption         =   "Graph Settings..."
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txtInput 
      Height          =   285
      Left            =   4680
      TabIndex        =   2
      Top             =   360
      Width           =   4935
   End
   Begin VB.CommandButton cmdRecip 
      Caption         =   "1 / x"
      Height          =   375
      Left            =   4680
      TabIndex        =   41
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      TabIndex        =   4
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton cmdFactorial 
      Caption         =   "x!"
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton cmdEPlus 
      Caption         =   "E+"
      Height          =   375
      Left            =   5520
      TabIndex        =   6
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton cmdENeg 
      Caption         =   "E-"
      Height          =   375
      Left            =   6360
      TabIndex        =   7
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton cmdLog 
      Caption         =   "Log"
      Height          =   375
      Left            =   7200
      TabIndex        =   8
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton cmdLn 
      Caption         =   "Ln"
      Height          =   375
      Left            =   8040
      TabIndex        =   9
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton cmdSin 
      Caption         =   "Sin"
      Height          =   375
      Left            =   4680
      TabIndex        =   11
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton cmdCos 
      Caption         =   "Cos"
      Height          =   375
      Left            =   6360
      TabIndex        =   13
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton cmdTan 
      Caption         =   "Tan"
      Height          =   375
      Left            =   8040
      TabIndex        =   15
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton cmdAsn 
      Caption         =   "Sin^-1"
      Height          =   375
      Left            =   5520
      TabIndex        =   12
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton cmdAcs 
      Caption         =   "Cos^-1"
      Height          =   375
      Left            =   7200
      TabIndex        =   14
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton cmdAtn 
      Caption         =   "Tan^-1"
      Height          =   375
      Left            =   8880
      TabIndex        =   16
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton cmdXtoYpower 
      Caption         =   "x^y"
      Height          =   375
      Left            =   8040
      TabIndex        =   21
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton cmdGetBase 
      Caption         =   "x^1/y"
      Height          =   375
      Left            =   5520
      TabIndex        =   18
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton cmdSqr 
      Caption         =   "Sqr"
      Height          =   375
      Left            =   4680
      TabIndex        =   17
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton cmdSquared 
      Caption         =   "x^2"
      Height          =   375
      Left            =   6360
      TabIndex        =   19
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton cmdNeg1Power 
      Caption         =   "x^-1"
      Height          =   375
      Left            =   7200
      TabIndex        =   20
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton cmdExp 
      Caption         =   "e^x"
      Height          =   375
      Left            =   8880
      TabIndex        =   10
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton cmdBackspace 
      Caption         =   "Bksp"
      Height          =   375
      Left            =   8880
      TabIndex        =   22
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton cmdOpenPar 
      Caption         =   "("
      Height          =   375
      Left            =   4680
      TabIndex        =   23
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton cmdClosedPar 
      Caption         =   ")"
      Height          =   375
      Left            =   4680
      TabIndex        =   29
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton cmdNum7 
      Caption         =   "7"
      Height          =   375
      Left            =   5520
      TabIndex        =   24
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton cmdNum8 
      Caption         =   "8"
      Height          =   375
      Left            =   6360
      TabIndex        =   25
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton cmdNum9 
      Caption         =   "9"
      Height          =   375
      Left            =   7200
      TabIndex        =   26
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton cmdMinusOrNeg 
      Caption         =   "-"
      Height          =   375
      Left            =   8040
      TabIndex        =   33
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton cmdTimes 
      Caption         =   "×"
      Height          =   375
      Left            =   8880
      TabIndex        =   34
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton cmdDividedBy 
      Caption         =   "÷"
      Height          =   375
      Left            =   8880
      TabIndex        =   40
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton cmdPlus 
      Caption         =   "+"
      Height          =   375
      Left            =   8040
      TabIndex        =   39
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton cmdPi 
      Caption         =   "Pi"
      Height          =   375
      Left            =   7200
      TabIndex        =   44
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton cmdGraph 
      Caption         =   "Graph"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   45
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton cmdNum4 
      Caption         =   "4"
      Height          =   375
      Left            =   5520
      TabIndex        =   30
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton cmdNum5 
      Caption         =   "5"
      Height          =   375
      Left            =   6360
      TabIndex        =   31
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton cmdNum6 
      Caption         =   "6"
      Height          =   375
      Left            =   7200
      TabIndex        =   32
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton cmdNum1 
      Caption         =   "1"
      Height          =   375
      Left            =   5520
      TabIndex        =   36
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton cmdNum2 
      Caption         =   "2"
      Height          =   375
      Left            =   6360
      TabIndex        =   37
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton cmdNum3 
      Caption         =   "3"
      Height          =   375
      Left            =   7200
      TabIndex        =   38
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton cmdNum0 
      Caption         =   "0"
      Height          =   375
      Left            =   5520
      TabIndex        =   42
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton cmdDecPoint 
      Caption         =   "."
      Height          =   375
      Left            =   6360
      TabIndex        =   43
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton cmdXVar 
      Caption         =   "x"
      Height          =   375
      Left            =   4680
      TabIndex        =   35
      Top             =   3600
      Width           =   735
   End
   Begin VB.Timer tmrGrapher 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox pbGraph 
      Height          =   4575
      Left            =   0
      ScaleHeight     =   4515
      ScaleWidth      =   4515
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
   Begin VB.Label lblInput 
      Caption         =   "Enter f(x) here, and only use ""x"" as your variable."
      Height          =   255
      Left            =   4680
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim it As Long
Dim TheExp As String
Dim bLastError As Boolean
Dim LastY As Double

Private Function TP(ByVal Tw As Single) As Single
TP = Tw / Screen.TwipsPerPixelX
End Function

Private Function PurelyNumeric(Expression) As Boolean
PurelyNumeric = IsNumeric(Expression) And (InStr(CStr(Expression), "(") = 0 And InStr(CStr(Expression), ")") = 0)
End Function

Private Function IsNegative(ByVal Number As Double) As Boolean
IsNegative = (Number < 0)
End Function

Private Function Opposite(ByVal Number As Double) As Double
Opposite = -Number
End Function

Private Function DoCount(ByVal sInput As String, ByVal sMatch As String, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare) As Long
Dim lLen As Long
Dim i As Long
Dim j As Long

lLen = Len(sMatch)
If lLen = 0 Then Exit Function

Do
    i = InStr(i + lLen, sInput, sMatch, Compare)
    If i > 0 Then
        j = j + 1
    Else
        Exit Do
    End If
Loop

DoCount = j
End Function

Private Function RangeIt(ByVal Exp As String, ByVal Start As Long, InPercent As Boolean, InFactorial As Boolean, Optional Reverse As Boolean = False, Optional InPar As Boolean) As Long
On Error Resume Next

Dim tStart As Long
Dim X As String
Dim Y As String
Dim bSkip As Boolean
Dim MyExpr As String

bSkip = False
InPercent = False
InFactorial = False
InPar = False
tStart = Start

Do
    tStart = tStart + IIf(Reverse, -1, 1)
    If tStart > Len(Exp) Or tStart <= 0 Then Exit Do
    
    If bSkip Then
        bSkip = False
    Else
        X = UCase$(Mid$(Exp, tStart, 1))
        Select Case X
            Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", "%", "!", "(", ")", ","
            Case "E"
                If Not Reverse Then
                    Y = Mid$(Exp, tStart + 1, 1)
                    If Y = "+" Or Y = "-" Then
                        bSkip = True
                    Else
                        Exit Do
                    End If
                End If
            Case "-"
                If Reverse Then
                    Y = UCase$(Mid$(Exp, tStart - 1, 1))
                    Select Case Y
                        Case "+", "-", "*", "/", "": tStart = tStart - 1: Exit Do
                        Case "(": tStart = tStart - 2: Exit Do
                        Case "E": bSkip = True
                        Case Else: Exit Do
                    End Select
                Else
                    If tStart - Start > 1 Then Exit Do
                End If
            Case "+"
                If Reverse Then
                    Y = UCase$(Mid$(Exp, tStart - 1, 1))
                    If Y = "E" Then
                        bSkip = True
                    Else
                        Exit Do
                    End If
                Else
                    Exit Do
                End If
            Case Else: Exit Do
        End Select
    End If
Loop

tStart = tStart + IIf(Reverse, 1, -1)

If Reverse Then
    MyExpr = Mid$(Exp, tStart, Start - tStart)
Else
    MyExpr = Mid$(Exp, Start + 1, tStart - Start)
End If

InPercent = (Right$(Replace(MyExpr, "!", ""), 1) = "%")
InFactorial = (Right$(Replace(MyExpr, "%", ""), 1) = "!")
InPar = ((Left$(MyExpr, 1) = "(") And (Right$(MyExpr, 1) = ")"))

RangeIt = tStart
End Function

'THE MOST IMPORTANT FUNCTION OF THE CALCULATOR!!!
Private Function EvaluateInput(ByVal sInput As String) As String
On Error GoTo ProcExit

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim m As Long
Dim tmpStr As String
Dim tmpStr2 As String
Dim tmpStr3 As String
Dim tmpExp As String
Dim tmpExp2 As String
Dim tmpNum As Double
Dim tmpNumerator As Double
Dim tmpDenominator As Double
Dim NoFraction As Boolean

Dim lOpenParCount As Long
Dim lClosedParCount As Long

Dim ColPar As Collection

Dim bPercent As Boolean
Dim bPercent2 As Boolean

Dim bFactorial As Boolean
Dim bFactorial2 As Boolean

Dim bPar As Boolean

Dim D1 As Double
Dim D2 As Double

NoFraction = True
tmpStr = LCase$(sInput)
tmpStr = Replace(tmpStr, " ", "")
tmpStr = Replace(tmpStr, "pi", Replace(CStr(Pi), ",", "."))

tmpStr2 = tmpStr

lOpenParCount = DoCount(tmpStr2, "(")
lClosedParCount = DoCount(tmpStr2, ")")

If lOpenParCount > lClosedParCount Then
    tmpStr2 = tmpStr2 & String$(lOpenParCount - lClosedParCount, ")")
ElseIf lClosedParCount > lOpenParCount Then
    EvaluateInput = "Syntax error"
    Exit Function
End If

Do
    i = InStr(k + 1, tmpStr2, "(")
    j = InStr(k + 1, tmpStr2, ")")
    
    If i > 1 Or j > 1 Then
        If i <= 1 Or j < i Then
            k = j
            If PurelyNumeric(Mid$(tmpStr2, k + 1, 1)) Or (Mid$(tmpStr2, k + 1, 1) = "x") Or (Mid$(tmpStr2, k + 1, 1) = "(") Then
                tmpStr2 = Left$(tmpStr2, k) & "*" & Mid$(tmpStr2, k + 1)
            End If
        Else
            k = i
            If PurelyNumeric(Mid$(tmpStr2, k - 1, 1)) Or (Mid$(tmpStr2, k - 1, 1) = "x") Or (Mid$(tmpStr2, k - 1, 1) = ")") Then
                tmpStr2 = Left$(tmpStr2, k - 1) & "*" & Mid$(tmpStr2, k)
            ElseIf Mid$(tmpStr2, k - 1, 1) = "-" Then
                If k > 2 Then
                    If Mid$(tmpStr2, k - 2, 1) = "*" Or Mid$(tmpStr2, k - 2, 1) = "/" Then
                        tmpStr2 = Left$(tmpStr2, k - 2) & "(-1*" & Mid$(tmpStr2, k, j - k + 1) & ")" & Mid$(tmpStr2, j + 1)
                    Else
                        GoTo Continuation
                    End If
                Else
Continuation:       tmpStr2 = Left$(tmpStr2, k - 1) & "1*" & Mid$(tmpStr2, k)
                End If
            End If
        End If
    Else
        If i = 0 And j = 0 Then Exit Do
    End If
Loop

RedoOperations:

If Not PurelyNumeric(tmpStr2) Then  'Expression, not a number
    
    ' Refresh
    
    i = 0
    j = 0
    
SignFixProc:
    tmpStr3 = tmpStr2
    Do
        i = InStr(i + 1, tmpStr2, "--")
        If i = 1 Then
            tmpStr2 = Mid$(tmpStr2, 3)
        Else
            If i > 0 Then
                tmpStr2 = Left$(tmpStr2, i - 1) & "+" & Mid$(tmpStr2, i + 2)
            Else
                Exit Do
            End If
        End If
    Loop
    
    Do
        i = InStr(i + 1, tmpStr2, "+-")
        If i = 1 Then
            tmpStr2 = Mid$(tmpStr2, 2)
        Else
            If i > 0 Then
                tmpStr2 = Left$(tmpStr2, i - 1) & Mid$(tmpStr2, i + 1)
            Else
                Exit Do
            End If
        End If
    Loop
    
    Do
        i = InStr(i + 1, tmpStr2, "-+")
        If i = 1 Then
            tmpStr2 = "-" & Mid$(tmpStr2, 3)
        Else
            If i > 0 Then
                tmpStr2 = Left$(tmpStr2, i) & Mid$(tmpStr2, i + 2)
            Else
                Exit Do
            End If
        End If
    Loop
    
    tmpStr2 = Replace(tmpStr2, "++", "+")
    
    If tmpStr3 <> tmpStr2 Then GoTo SignFixProc
    
    'Refresh
    
    i = 0
    j = 0
    tmpStr3 = tmpStr2
    
    Set ColPar = Nothing
    Set ColPar = New Collection
    
    'Do order of operations
    
    Do
        j = InStr(j + 1, tmpStr3, ")")
        If j > 0 Then
            i = InStrRev(tmpStr3, "(", j - 1)
            If i = 0 Then
                EvaluateInput = "Syntax error"
                Exit Function
            Else
                tmpStr3 = Left$(tmpStr3, i - 1) & "<" & Mid$(tmpStr3, i + 1)
                tmpStr3 = Left$(tmpStr3, j - 1) & ">" & Mid$(tmpStr3, j + 1)
                ColPar.Add j, "P" & CStr(i)
            End If
        Else
            Exit Do
        End If
    Loop
    
    i = InStr(tmpStr2, "log(")
    If i > 0 Then
        j = ColPar("P" & CStr(i + 3))
        If j > 0 Then
            tmpExp = EvaluateInput(Mid$(tmpStr2, i + 4, j - i - 4))
            If PurelyNumeric(tmpExp) Then
                tmpStr2 = Left$(tmpStr2, i - 1) & CStr(Log10(CDbl(tmpExp))) & Mid$(tmpStr2, j + 1)
            Else
                GoTo ProcExit
            End If
        Else
            GoTo ProcExit
        End If
        GoTo RedoOperations
    End If
    
    i = InStr(tmpStr2, "ln(")
    If i > 0 Then
        j = ColPar("P" & CStr(i + 2))
        If j > 0 Then
            tmpExp = EvaluateInput(Mid$(tmpStr2, i + 3, j - i - 3))
            If PurelyNumeric(tmpExp) Then
                tmpStr2 = Left$(tmpStr2, i - 1) & CStr(Log(CDbl(tmpExp))) & Mid$(tmpStr2, j + 1)
            Else
                GoTo ProcExit
            End If
        Else
            GoTo ProcExit
        End If
        GoTo RedoOperations
    End If
    
    i = InStr(tmpStr2, "exp(")
    If i > 0 Then
        j = ColPar("P" & CStr(i + 3))
        If j > 0 Then
            tmpExp = EvaluateInput(Mid$(tmpStr2, i + 4, j - i - 4))
            If PurelyNumeric(tmpExp) Then
                tmpStr2 = Left$(tmpStr2, i - 1) & CStr(Exp(CDbl(tmpExp))) & Mid$(tmpStr2, j + 1)
            Else
                GoTo ProcExit
            End If
        Else
            GoTo ProcExit
        End If
        GoTo RedoOperations
    End If
    
    i = InStr(tmpStr2, "sin(")
    If i > 0 Then
        j = ColPar("P" & CStr(i + 3))
        If j > 0 Then
            tmpExp = EvaluateInput(Mid$(tmpStr2, i + 4, j - i - 4))
            If PurelyNumeric(tmpExp) Then
                tmpNum = Sin(CDbl(tmpExp))
                tmpStr2 = Left$(tmpStr2, i - 1) & CStr(tmpNum) & Mid$(tmpStr2, j + 1)
            Else
                GoTo ProcExit
            End If
        Else
            GoTo ProcExit
        End If
        GoTo RedoOperations
    End If
    
    i = InStr(tmpStr2, "asn(")
    If i > 0 Then
        j = ColPar("P" & CStr(i + 3))
        If j > 0 Then
            tmpExp = EvaluateInput(Mid$(tmpStr2, i + 4, j - i - 4))
            If PurelyNumeric(tmpExp) Then
                tmpNum = Asn(CDbl(tmpExp))
                tmpStr2 = Left$(tmpStr2, i - 1) & CStr(tmpNum) & Mid$(tmpStr2, j + 1)
            Else
                GoTo ProcExit
            End If
        Else
            GoTo ProcExit
        End If
        GoTo RedoOperations
    End If
    
    i = InStr(tmpStr2, "cos(")
    If i > 0 Then
        j = ColPar("P" & CStr(i + 3))
        If j > 0 Then
            tmpExp = EvaluateInput(Mid$(tmpStr2, i + 4, j - i - 4))
            If PurelyNumeric(tmpExp) Then
                tmpNum = Cos(CDbl(tmpExp))
                tmpStr2 = Left$(tmpStr2, i - 1) & CStr(tmpNum) & Mid$(tmpStr2, j + 1)
            Else
                GoTo ProcExit
            End If
        Else
            GoTo ProcExit
        End If
        GoTo RedoOperations
    End If
    
    i = InStr(tmpStr2, "acs(")
    If i > 0 Then
        j = ColPar("P" & CStr(i + 3))
        If j > 0 Then
            tmpExp = EvaluateInput(Mid$(tmpStr2, i + 4, j - i - 4))
            If PurelyNumeric(tmpExp) Then
                tmpNum = Acs(CDbl(tmpExp))
                tmpStr2 = Left$(tmpStr2, i - 1) & CStr(tmpNum) & Mid$(tmpStr2, j + 1)
            Else
                GoTo ProcExit
            End If
        Else
            GoTo ProcExit
        End If
        GoTo RedoOperations
    End If
    
    i = InStr(tmpStr2, "tan(")
    If i > 0 Then
        j = ColPar("P" & CStr(i + 3))
        If j > 0 Then
            tmpExp = EvaluateInput(Mid$(tmpStr2, i + 4, j - i - 4))
            If PurelyNumeric(tmpExp) Then
                tmpNum = Tan(CDbl(tmpExp))
                tmpStr2 = Left$(tmpStr2, i - 1) & CStr(tmpNum) & Mid$(tmpStr2, j + 1)
            Else
                GoTo ProcExit
            End If
        Else
            GoTo ProcExit
        End If
        GoTo RedoOperations
    End If
    
    i = InStr(tmpStr2, "atn(")
    If i > 0 Then
        j = ColPar("P" & CStr(i + 3))
        If j > 0 Then
            tmpExp = EvaluateInput(Mid$(tmpStr2, i + 4, j - i - 4))
            If PurelyNumeric(tmpExp) Then
                tmpNum = Atn(CDbl(tmpExp))
                tmpStr2 = Left$(tmpStr2, i - 1) & CStr(tmpNum) & Mid$(tmpStr2, j + 1)
            Else
                GoTo ProcExit
            End If
        Else
            GoTo ProcExit
        End If
        GoTo RedoOperations
    End If
    
    i = InStrRev(tmpStr2, "sqr(")
    If i > 0 Then
        j = ColPar("P" & CStr(i + 3))
        If j > 0 Then
            tmpExp = EvaluateInput(Mid$(tmpStr2, i + 4, j - i - 4))
            If PurelyNumeric(tmpExp) Then
                tmpStr2 = Left$(tmpStr2, i - 1) & CStr(Sqr(CDbl(tmpExp))) & Mid$(tmpStr2, j + 1)
            Else
                GoTo ProcExit
            End If
        Else
            GoTo ProcExit
        End If
        GoTo RedoOperations
    End If
    
    i = InStr(tmpStr2, "int(")
    If i > 0 Then
        j = ColPar("P" & CStr(i + 3))
        If j > 0 Then
            tmpExp = EvaluateInput(Mid$(tmpStr2, i + 4, j - i - 4))
            If PurelyNumeric(tmpExp) Then
                tmpStr2 = Left$(tmpStr2, i - 1) & CStr(Fix(CDbl(tmpExp))) & Mid$(tmpStr2, j + 1)
            Else
                GoTo ProcExit
            End If
        Else
            GoTo ProcExit
        End If
        GoTo RedoOperations
    End If
    
    i = InStr(tmpStr2, "abs(")
    If i > 0 Then
        j = ColPar("P" & CStr(i + 3))
        If j > 0 Then
            tmpExp = EvaluateInput(Mid$(tmpStr2, i + 4, j - i - 4))
            If PurelyNumeric(tmpExp) Then
                tmpStr2 = Left$(tmpStr2, i - 1) & CStr(Abs(CDbl(tmpExp))) & Mid$(tmpStr2, j + 1)
            Else
                GoTo ProcExit
            End If
        Else
            GoTo ProcExit
        End If
        GoTo RedoOperations
    End If
    
    i = 0
ProcParCheck:
    i = InStr(i + 1, tmpStr2, "(")
    If i > 0 Then
        j = ColPar("P" & CStr(i))
        If j > 0 Then
            tmpExp = Mid$(tmpStr2, i + 1, j - i - 1)
            If PurelyNumeric(tmpExp) Then
                If Mid$(tmpStr2, j + 1, 1) = "^" Then
                    GoTo ProcParCheck
                Else
                    tmpExp2 = EvaluateInput(Mid$(tmpStr2, i + 1, j - i - 1))
                    If PurelyNumeric(tmpExp2) Then
                        tmpStr2 = Left$(tmpStr2, i - 1) & tmpExp2 & Mid$(tmpStr2, j + 1)
                    Else
                        GoTo ProcExit
                    End If
                End If
            Else
                tmpExp2 = EvaluateInput(Mid$(tmpStr2, i + 1, j - i - 1))
                If PurelyNumeric(tmpExp2) Then
                    tmpStr2 = Left$(tmpStr2, i) & tmpExp2 & Mid$(tmpStr2, j)
                Else
                    GoTo ProcExit
                End If
            End If
        Else
            GoTo ProcExit
        End If
        GoTo RedoOperations
    End If
    
    i = InStrRev(tmpStr2, "^")
    If i > 0 Then
        j = RangeIt(tmpStr2, i, bPercent, bFactorial)
        k = RangeIt(tmpStr2, i, bPercent2, bFactorial2, True, bPar)
        
        tmpExp = Replace(Replace(Replace(Replace(Mid$(tmpStr2, k, i - k), "%", ""), "!", ""), "(", ""), ")", "")
        If PurelyNumeric(tmpExp) Then
            D1 = CDbl(tmpExp)
            If bPercent2 Then D1 = D1 / 100
            If bFactorial2 Then D1 = Factorial(D1)
        Else
            GoTo ProcExit
        End If
        
        tmpExp = Replace(Replace(Replace(Replace(Mid$(tmpStr2, i + 1, j - i), "%", ""), "!", ""), "(", ""), ")", "")
        If PurelyNumeric(tmpExp) Then
            D2 = CDbl(tmpExp)
            If bPercent Then D2 = D2 / 100
            If bFactorial Then D2 = Factorial(D2)
        Else
            GoTo ProcExit
        End If
        
        If bPar Then
            tmpNum = D1 ^ D2
        Else
            tmpNum = Abs(D1) ^ D2
            If IsNegative(D1) Then tmpNum = Opposite(tmpNum)
        End If
        
        tmpStr2 = Left$(tmpStr2, k - 1) & CStr(tmpNum) & Mid$(tmpStr2, j + 1)
        GoTo RedoOperations
    End If
    
    i = InStr(tmpStr2, "*")
    j = InStr(tmpStr2, "/")
    k = InStr(tmpStr2, "mod")
    
    If i > 0 And j = 0 And k = 0 Then GoTo MulProc
    If j > 0 And i = 0 And k = 0 Then GoTo DivProc
    If k > 0 And i = 0 And j = 0 Then GoTo ModProc
    
    If ((i > 0 And j > 0 And k = 0) And j > i) Or ((i > 0 And k > 0 And j = 0) And k > i) Or ((i > 0 And j > 0 And k > 0) And j > i And k > i) Then
MulProc:
        l = RangeIt(tmpStr2, i, bPercent, bFactorial)
        m = RangeIt(tmpStr2, i, bPercent2, bFactorial2, True)
        
        tmpExp = Replace(Replace(Mid$(tmpStr2, m, i - m), "%", ""), "!", "")
        If PurelyNumeric(tmpExp) Then
            D1 = CDbl(tmpExp)
            If bFactorial2 Then D1 = Factorial(D1)
            If bPercent2 Then D1 = D1 / 100
        Else
            GoTo ProcExit
        End If
        
        tmpExp = Replace(Replace(Mid$(tmpStr2, i + 1, l - i), "%", ""), "!", "")
        If PurelyNumeric(tmpExp) Then
            D2 = CDbl(tmpExp)
            If bFactorial Then D2 = Factorial(D2)
            If bPercent Then D2 = D2 / 100
        Else
            GoTo ProcExit
        End If
        
        tmpStr2 = Left$(tmpStr2, m - 1) & CStr(D1 * D2) & Mid$(tmpStr2, l + 1)
        GoTo RedoOperations
    ElseIf ((i > 0 And j > 0 And k = 0) And i > j) Or ((j > 0 And k > 0 And i = 0) And k > j) Or ((i > 0 And j > 0 And k > 0) And i > j And k > j) Then
DivProc:
        l = RangeIt(tmpStr2, j, bPercent, bFactorial)
        m = RangeIt(tmpStr2, j, bPercent2, bFactorial2, True)
        
        tmpExp = Replace(Replace(Mid$(tmpStr2, m, j - m), "%", ""), "!", "")
        If PurelyNumeric(tmpExp) Then
            D1 = CDbl(tmpExp)
            If bFactorial2 Then D1 = Factorial(D1)
            If bPercent2 Then D1 = D1 / 100
        Else
            GoTo ProcExit
        End If
        
        tmpExp = Replace(Replace(Mid$(tmpStr2, j + 1, l - j), "%", ""), "!", "")
        If PurelyNumeric(tmpExp) Then
            D2 = CDbl(tmpExp)
            If bFactorial Then D2 = Factorial(D2)
            If bPercent Then D2 = D2 / 100
        Else
            GoTo ProcExit
        End If
        
        tmpStr2 = Left$(tmpStr2, m - 1) & CStr(D1 / D2) & Mid$(tmpStr2, l + 1)
        GoTo RedoOperations
    ElseIf ((i > 0 And k > 0 And j = 0) And i > k) Or ((j > 0 And k > 0 And i = 0) And j > k) Or ((i > 0 And j > 0 And k > 0) And i > k And j > k) Then
ModProc:
        l = RangeIt(tmpStr2, k + 2, bPercent, bFactorial)
        m = RangeIt(tmpStr2, k, bPercent2, bFactorial2, True)
        
        tmpExp = Replace(Replace(Mid$(tmpStr2, m, k - m), "%", ""), "!", "")
        If PurelyNumeric(tmpExp) Then
            D1 = CDbl(tmpExp)
            If bFactorial2 Then D1 = Factorial(D1)
            If bPercent2 Then D1 = D1 / 100
        Else
            GoTo ProcExit
        End If
        
        tmpExp = Replace(Replace(Mid$(tmpStr2, k + 3, l - k - 2), "%", ""), "!", "")
        If PurelyNumeric(tmpExp) Then
            D2 = CDbl(tmpExp)
            If bFactorial Then D2 = Factorial(D2)
            If bPercent Then D2 = D2 / 100
        Else
            GoTo ProcExit
        End If
        
        tmpStr2 = Left$(tmpStr2, m - 1) & CStr(Modulo(D1, D2)) & Mid$(tmpStr2, l + 1)
        GoTo RedoOperations
    End If
    
    i = InStr(tmpStr2, "+")
    j = InStr(tmpStr2, "-")
    If j = 1 Then j = InStr(2, tmpStr2, "-")
    
AddSubProc:
    
    If i > 0 And j = 0 Then GoTo AddProc
    If j > 0 And i = 0 Then GoTo SubProc
    
    If (i > 0 Or j > 0) And j > i Then
AddProc:
        If Mid$(tmpStr2, i - 1, 1) <> "E" Then
            k = RangeIt(tmpStr2, i, bPercent, bFactorial)
            l = RangeIt(tmpStr2, i, bPercent2, bFactorial2, True)
            
            tmpExp = Replace(Replace(Mid$(tmpStr2, l, i - 1), "%", ""), "!", "")
            If PurelyNumeric(tmpExp) Then
                D1 = CDbl(tmpExp)
                If bFactorial2 Then D1 = Factorial(D1)
                If bPercent2 Then D1 = D1 / 100
            Else
                GoTo ProcExit
            End If
            
            tmpExp = Replace(Replace(Mid$(tmpStr2, i + 1, k - i), "%", ""), "!", "")
            If PurelyNumeric(tmpExp) Then
                D2 = CDbl(tmpExp)
                If bFactorial Then D2 = Factorial(D2)
                If bPercent Then D2 = D2 / 100
            Else
                GoTo ProcExit
            End If
            
            tmpStr2 = Left$(tmpStr2, l - 1) & CStr(D1 + D2) & Mid$(tmpStr2, k + 1)
            GoTo RedoOperations
        Else
            i = InStr(i + 1, tmpStr2, "-")
            GoTo AddSubProc
        End If
    ElseIf (i > 0 Or j > 0) And i > j Then
SubProc:
        If Mid$(tmpStr2, j - 1, 1) <> "E" Then
            k = RangeIt(tmpStr2, j, bPercent, bFactorial)
            l = RangeIt(tmpStr2, j, bPercent2, bFactorial2, True)
            
            tmpExp = Replace(Replace(Mid$(tmpStr2, l, j - 1), "%", ""), "!", "")
            If PurelyNumeric(tmpExp) Then
                D1 = CDbl(tmpExp)
                If bFactorial2 Then D1 = Factorial(D1)
                If bPercent2 Then D1 = D1 / 100
            Else
                GoTo ProcExit
            End If
            
            tmpExp = Replace(Replace(Mid$(tmpStr2, j + 1, k - j), "%", ""), "!", "")
            If PurelyNumeric(tmpExp) Then
                D2 = CDbl(tmpExp)
                If bFactorial Then D2 = Factorial(D2)
                If bPercent Then D2 = D2 / 100
            Else
                GoTo ProcExit
            End If
            
            tmpStr2 = Left$(tmpStr2, l - 1) & CStr(D1 - D2) & Mid$(tmpStr2, k + 1)
            GoTo RedoOperations
        Else
            j = InStr(j + 1, tmpStr2, "-")
            GoTo AddSubProc
        End If
    End If
    
    bPercent = (Right$(Replace(tmpStr2, "!", ""), 1) = "%")
    bFactorial = (Right$(Replace(tmpStr2, "%", ""), 1) = "!")
    
    If bFactorial Then tmpStr2 = CStr(Factorial(CDbl(Replace(Replace(tmpStr2, "%", ""), "!", ""))))
    If bPercent Then tmpStr2 = CStr(CDbl(Replace(Replace(tmpStr2, "%", ""), "!", "")) / 100)
Else
    tmpStr2 = CStr(CDbl(tmpStr2))
End If

If Not PurelyNumeric(tmpStr2) Then
    EvaluateInput = "Error"
    Exit Function
End If

EvaluateInput = tmpStr2

Exit Function
ProcExit:
    EvaluateInput = "Error"
End Function

Private Sub AddTextToInput(ByVal Text As String)
txtInput = txtInput & Text
End Sub

Private Sub FocusAndSetCursor()
txtInput.SetFocus
txtInput.SelStart = Len(txtInput)
End Sub

Private Sub AddTextAndFocus(ByVal Text As String)
AddTextToInput Text
FocusAndSetCursor
End Sub

Private Sub cmdAbs_Click()
AddTextAndFocus "Abs("
End Sub

Private Sub cmdAcs_Click()
AddTextAndFocus "Acs("
End Sub

Private Sub cmdAsn_Click()
AddTextAndFocus "Asn("
End Sub

Private Sub cmdAtn_Click()
AddTextAndFocus "Atn("
End Sub

Private Sub cmdBackspace_Click()
On Error Resume Next
If txtInput <> "" Then txtInput = Left$(txtInput, Len(txtInput) - 1)
FocusAndSetCursor
End Sub

Private Sub cmdClear_Click()
txtInput = ""
FocusAndSetCursor
End Sub

Private Sub cmdClosedPar_Click()
AddTextAndFocus ")"
End Sub

Private Sub cmdCos_Click()
AddTextAndFocus "Cos("
End Sub

Private Sub cmdDecPoint_Click()
AddTextAndFocus "."
End Sub

Private Sub cmdDividedBy_Click()
AddTextAndFocus "/"
End Sub

Private Sub cmdENeg_Click()
AddTextAndFocus "E-"
End Sub

Private Sub cmdEPlus_Click()
AddTextAndFocus "E+"
End Sub

Private Sub cmdGraph_Click()
On Error Resume Next
Dim i As Long

If txtInput <> "" Then
    MousePointer = vbHourglass: DoEvents
    
    Set pbGraph.Picture = Nothing
    pbGraph.AutoRedraw = True
    PaintLines pbGraph, XMin, XMax, YMin, YMax, XScl, YScl
    
    it = 0
    pbGraph.CurrentX = 0
    pbGraph.DrawWidth = 2
    TheExp = LCase$(txtInput)
    
    Do
        i = InStr(i + 1, TheExp, "x")
        If i > 1 Then
            If PurelyNumeric(Mid$(TheExp, i - 1, 1)) Or (Mid$(TheExp, i - 1, 1) = "x") Then
                TheExp = Left$(TheExp, i - 1) & "*" & Mid$(TheExp, i)
                i = i + 1
            End If
        Else
            If i = 0 Then Exit Do
        End If
    Loop
    
    bLastError = True
    LastY = 0
    tmrGrapher.Enabled = True
    tmrGrapher_Timer
End If
End Sub

Private Sub cmdExp_Click()
AddTextAndFocus "Exp("
End Sub

Private Sub cmdFactorial_Click()
AddTextAndFocus "!"
End Sub

Private Sub cmdGetBase_Click()
AddTextAndFocus "^(1/("
End Sub

Private Sub cmdGraphSettings_Click()
frmGraphSettings.Show 1
End Sub

Private Sub cmdLn_Click()
AddTextAndFocus "Ln("
End Sub

Private Sub cmdLog_Click()
AddTextAndFocus "Log("
End Sub

Private Sub cmdMinusOrNeg_Click()
AddTextAndFocus "-"
End Sub

Private Sub cmdMod_Click()
AddTextAndFocus " Mod "
End Sub

Private Sub cmdNeg1Power_Click()
AddTextAndFocus "^-1"
End Sub

Private Sub cmdNum0_Click()
AddTextAndFocus "0"
End Sub

Private Sub cmdNum1_Click()
AddTextAndFocus "1"
End Sub

Private Sub cmdNum2_Click()
AddTextAndFocus "2"
End Sub

Private Sub cmdNum3_Click()
AddTextAndFocus "3"
End Sub

Private Sub cmdNum4_Click()
AddTextAndFocus "4"
End Sub

Private Sub cmdNum5_Click()
AddTextAndFocus "5"
End Sub

Private Sub cmdNum6_Click()
AddTextAndFocus "6"
End Sub

Private Sub cmdNum7_Click()
AddTextAndFocus "7"
End Sub

Private Sub cmdNum8_Click()
AddTextAndFocus "8"
End Sub

Private Sub cmdNum9_Click()
AddTextAndFocus "9"
End Sub

Private Sub cmdOpenPar_Click()
AddTextAndFocus "("
End Sub

Private Sub cmdPi_Click()
AddTextAndFocus "Pi"
End Sub

Private Sub cmdPlus_Click()
AddTextAndFocus "+"
End Sub

Private Sub cmdRecip_Click()
AddTextAndFocus "1/("
End Sub

Private Sub cmdSin_Click()
AddTextAndFocus "Sin("
End Sub

Private Sub cmdSqr_Click()
AddTextAndFocus "Sqr("
End Sub

Private Sub cmdSquared_Click()
AddTextAndFocus "^2"
End Sub

Private Sub cmdTan_Click()
AddTextAndFocus "Tan("
End Sub

Private Sub cmdTimes_Click()
AddTextAndFocus "*"
End Sub

Private Sub cmdXtoYpower_Click()
AddTextAndFocus "^("
End Sub

Private Sub cmdXVar_Click()
AddTextAndFocus "x"
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveSetting App.Title, "GraphConfig", "XMin", XMin
SaveSetting App.Title, "GraphConfig", "XMax", XMax
SaveSetting App.Title, "GraphConfig", "YMin", YMin
SaveSetting App.Title, "GraphConfig", "YMax", YMax
SaveSetting App.Title, "GraphConfig", "XScale", XScl
SaveSetting App.Title, "GraphConfig", "YScale", YScl
End Sub

Private Sub tmrGrapher_Timer()
On Error Resume Next

Dim itScaled As Double
Dim i As Long
Dim X As String
Dim Y As Double
Dim tExp As String

tExp = TheExp
itScaled = ToXOfScale2(it, TP(60), TP(pbGraph.ScaleWidth - 60), XMin, XMax)
X = "(" & CStr(itScaled) & ")"

Do
    i = InStr(i + 1, tExp, "x")
    If i > 1 And i < Len(tExp) Then
        If Mid$(tExp, i - 1, 1) <> "e" And Mid$(tExp, i + 1, 1) <> "p" Then
            tExp = Left$(tExp, i - 1) & X & Mid$(tExp, i + 1)
            i = i + Len(X)
        End If
    Else
        If i = 0 Then
            Exit Do
        Else
            tExp = Left$(tExp, i - 1) & X & Mid$(tExp, i + 1)
            i = i + Len(X)
        End If
    End If
Loop

If Err.Number <> 0 Then Err.Clear
Y = EvaluateInput(tExp)
If Err.Number <> 0 And Err.Number <> 13 Then Err.Clear
Y = ToXOfScale2(Y, YMin, YMax, pbGraph.ScaleHeight - 60, 60)

If it = 0 Then
    pbGraph.CurrentX = it * Screen.TwipsPerPixelX
    pbGraph.CurrentY = Y
End If

If Err.Number <> 0 Then
    If Not bLastError Then
        bLastError = True
        pbGraph.CurrentX = it * Screen.TwipsPerPixelX
        pbGraph.CurrentY = LastY
    End If
    Err.Clear
Else
    If bLastError Then
        bLastError = False
        pbGraph.CurrentX = it * Screen.TwipsPerPixelX
        pbGraph.CurrentY = Y
    End If
    
    pbGraph.Line -(it * Screen.TwipsPerPixelX, Y), &HC0
    LastY = Y
End If

it = it + 1
If it > TP(pbGraph.ScaleWidth) Then
    TheExp = ""
    it = 0
    pbGraph.AutoRedraw = False
    tmrGrapher.Enabled = False
    MousePointer = vbNormal: DoEvents
End If
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
If txtInput <> "" Then
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmdGraph_Click
    End If
End If
End Sub

Private Function ToXOfScale2(ByVal XOfScale1 As Double, ByVal Scale1_Min As Double, ByVal Scale1_Max As Double, ByVal Scale2_Min As Double, ByVal Scale2_Max As Double) As Double
ToXOfScale2 = Scale2_Min + (Scale2_Max - Scale2_Min) * ((XOfScale1 - Scale1_Min) / (Scale1_Max - Scale1_Min))
End Function

Private Sub PaintLines(PBox As PictureBox, Optional ByVal XMin As Double = -10, Optional ByVal XMax As Double = 10, Optional ByVal YMin As Double = -10, Optional ByVal YMax As Double = 10, Optional ByVal XScale As Double = 1, Optional ByVal YScale As Double = 1)
Dim i As Long
Dim i2 As Long
Dim bRedraw As Boolean

Dim X As Double
Dim Y As Double
Dim Z As Double

Dim Alpha As Double
Dim Beta As Double

Dim clr As Long

bRedraw = PBox.AutoRedraw
PBox.AutoRedraw = True

clr = &HC0C0C0
PBox.DrawWidth = 1

Y = PBox.ScaleHeight / 2 - ScaleX(60, ScaleMode, vbTwips)
Z = PBox.ScaleHeight / 2 + ScaleX(60, ScaleMode, vbTwips)

For i = XMin To XMax
    If i <> 0 Then
        X = ToXOfScale2(i, XMin, XMax, 60, PBox.ScaleWidth - 60)
        
        PBox.CurrentX = X
        If (i / XScale) = Fix(i / XScale) Then
            Alpha = 0
            Beta = PBox.ScaleHeight
        Else
            Alpha = Y
            Beta = Z
        End If
        
        PBox.CurrentY = Alpha
        PBox.Line -(X, Alpha), clr
        PBox.Line -(X, Beta), clr
        PBox.CurrentY = Beta
    End If
Next

Y = PBox.ScaleWidth / 2 - ScaleY(60, ScaleMode, vbTwips)
Z = PBox.ScaleWidth / 2 + ScaleY(60, ScaleMode, vbTwips)

For i = YMin To YMax
    If i <> 0 Then
        X = ToXOfScale2(i, YMin, YMax, 60, PBox.ScaleHeight - 60)
        
        PBox.CurrentY = X
        If (i / YScale) = Fix(i / YScale) Then
            Alpha = 0
            Beta = PBox.ScaleWidth
        Else
            Alpha = Y
            Beta = Z
        End If
        
        PBox.CurrentX = Alpha
        PBox.Line -(Alpha, X), clr
        PBox.Line -(Beta, X), clr
        PBox.CurrentX = Beta
    End If
Next

clr = 0
PBox.DrawWidth = 4

If XMin <= 0 And XMax >= 0 Then
    X = ToXOfScale2(0, XMin, XMax, 60, PBox.ScaleWidth - 60)
    PBox.CurrentX = X
    PBox.CurrentY = 0
    PBox.Line -(X, 0), clr
    PBox.Line -(X, PBox.ScaleHeight), clr
    PBox.CurrentY = PBox.ScaleHeight
End If

If YMin <= 0 And YMax >= 0 Then
    X = ToXOfScale2(0, YMin, YMax, 60, PBox.ScaleHeight - 60)
    PBox.CurrentY = X
    PBox.CurrentX = 0
    PBox.Line -(0, X), clr
    PBox.Line -(PBox.ScaleWidth, X), clr
    PBox.CurrentX = PBox.ScaleWidth
End If

PBox.AutoRedraw = bRedraw
End Sub

Private Sub Form_Load()
Dim S1
Dim S2
Dim S3
Dim S4
Dim S5
Dim S6

S1 = GetSetting(App.Title, "GraphConfig", "XMin", -10)
If PurelyNumeric(S1) Then
    If S1 > 1000 Or S1 < -1000 Then
        S1 = -10
    Else
        S1 = Fix(S1)
    End If
Else
    S1 = -10
End If

S2 = GetSetting(App.Title, "GraphConfig", "XMax", 10)
If PurelyNumeric(S2) Then
    If S2 > 1000 Or S2 < -1000 Then
        S2 = 10
    Else
        S2 = Fix(S2)
    End If
Else
    S2 = 10
End If

If S1 >= S2 Then
    S1 = -10
    S2 = 10
End If

S3 = GetSetting(App.Title, "GraphConfig", "YMin", -10)
If PurelyNumeric(S3) Then
    If S3 > 1000 Or S3 < -1000 Then
        S3 = -10
    Else
        S3 = Fix(S3)
    End If
Else
    S3 = -10
End If

S4 = GetSetting(App.Title, "GraphConfig", "YMax", 10)
If PurelyNumeric(S4) Then
    If S4 > 1000 Or S4 < -1000 Then
        S4 = 10
    Else
        S4 = Fix(S4)
    End If
Else
    S4 = 10
End If

If S3 >= S4 Then
    S3 = -10
    S4 = 10
End If

S5 = GetSetting(App.Title, "GraphConfig", "XScale", 1)
If PurelyNumeric(S5) Then
    If S5 > 1000 Or S5 < 1 Then
        S5 = 1
    Else
        S5 = Fix(S5)
    End If
Else
    S5 = 1
End If

S6 = GetSetting(App.Title, "GraphConfig", "YScale", 1)
If PurelyNumeric(S6) Then
    If S6 > 1000 Or S6 < 1 Then
        S6 = 1
    Else
        S6 = Fix(S6)
    End If
Else
    S6 = 1
End If

XMin = S1
XMax = S2
YMin = S3
YMax = S4
XScl = S5
YScl = S6

PaintLines pbGraph, S1, S2, S3, S4, S5, S6
End Sub

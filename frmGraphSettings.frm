VERSION 5.00
Begin VB.Form frmGraphSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Graph Settings"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2520
   Icon            =   "frmGraphSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   2520
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtYScale 
      Height          =   285
      Left            =   1080
      TabIndex        =   11
      Text            =   "1"
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1320
      TabIndex        =   14
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdGoDefaults 
      Caption         =   "Revert to Defaults"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox txtXScale 
      Height          =   285
      Left            =   1080
      TabIndex        =   9
      Text            =   "1"
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox txtYMaximum 
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      Text            =   "10"
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox txtYMinimum 
      Height          =   285
      Left            =   1080
      TabIndex        =   5
      Text            =   "-10"
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox txtXMaximum 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Text            =   "10"
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtXMinimum 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Text            =   "-10"
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblYScale 
      Alignment       =   1  'Right Justify
      Caption         =   "Y-Scale:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label lblXScale 
      Alignment       =   1  'Right Justify
      Caption         =   "X-Scale:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label lblYMaximum 
      Alignment       =   1  'Right Justify
      Caption         =   "Y-maximum:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblYMinimum 
      Alignment       =   1  'Right Justify
      Caption         =   "Y-minimum:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lblXMaximum 
      Alignment       =   1  'Right Justify
      Caption         =   "X-maximum:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblXMinimum 
      Alignment       =   1  'Right Justify
      Caption         =   "X-minimum:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmGraphSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub HighlightTBox(TBox As TextBox)
On Error Resume Next
TBox.SetFocus
TBox.SelStart = 0
TBox.SelLength = Len(TBox.Text)
End Sub

Private Sub CheckTBoxes()
If IsNumeric(txtXMinimum) Then
    If txtXMinimum > 1000 Or txtXMinimum < -1000 Then
        txtXMinimum = -10
    Else
        If txtXMinimum <> Fix(txtXMinimum) Then txtXMinimum = Fix(txtXMinimum)
    End If
End If

If IsNumeric(txtXMaximum) Then
    If txtXMaximum > 1000 Or txtXMaximum < -1000 Then
        txtXMaximum = 10
    Else
        If txtXMaximum <> Fix(txtXMaximum) Then txtXMaximum = Fix(txtXMaximum)
    End If
End If

If IsNumeric(txtYMinimum) Then
    If txtYMinimum > 1000 Or txtYMinimum < -1000 Then
        txtYMinimum = -10
    Else
        If txtYMinimum <> Fix(txtYMinimum) Then txtYMinimum = Fix(txtYMinimum)
    End If
End If

If IsNumeric(txtYMaximum) Then
    If txtYMaximum > 1000 Or txtYMaximum < -1000 Then
        txtYMaximum = 10
    Else
        If txtYMaximum <> Fix(txtYMaximum) Then txtYMaximum = Fix(txtYMaximum)
    End If
End If

If IsNumeric(txtXMinimum) And IsNumeric(txtXMaximum) Then
    If txtXMinimum >= txtXMaximum Then
        txtXMaximum = CStr(CDbl(txtXMinimum) + 1)
        If txtXMaximum = 1001 Then
            txtXMinimum = 999
            txtXMaximum = 1000
        End If
    End If
End If

If IsNumeric(txtYMinimum) And IsNumeric(txtYMaximum) Then
    If txtYMinimum >= txtYMaximum Then
        txtYMaximum = CStr(CDbl(txtYMinimum) + 1)
        If txtYMaximum = 1001 Then
            txtYMinimum = 999
            txtYMaximum = 1000
        End If
    End If
End If

If IsNumeric(txtXScale) Then
    If txtXScale > 1000 Or txtXScale < 1 Then
        txtXScale = 1
    Else
        If txtXScale <> Fix(txtXScale) Then txtXScale = Fix(txtXScale)
    End If
End If

If IsNumeric(txtYScale) Then
    If txtYScale > 1000 Or txtYScale < 1 Then
        txtYScale = 1
    Else
        If txtYScale <> Fix(txtYScale) Then txtYScale = Fix(txtYScale)
    End If
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdGoDefaults_Click()
txtXMinimum = -10
txtXMaximum = 10
txtYMinimum = -10
txtYMaximum = 10
txtXScale = 1
txtYScale = 1
End Sub

Private Sub cmdOK_Click()
CheckTBoxes

If Not IsNumeric(txtXMinimum) Then HighlightTBox txtXMinimum: Exit Sub
If Not IsNumeric(txtXMaximum) Then HighlightTBox txtXMaximum: Exit Sub
If Not IsNumeric(txtYMinimum) Then HighlightTBox txtYMinimum: Exit Sub
If Not IsNumeric(txtYMaximum) Then HighlightTBox txtYMaximum: Exit Sub
If Not IsNumeric(txtXScale) Then HighlightTBox txtXScale: Exit Sub
If Not IsNumeric(txtYScale) Then HighlightTBox txtYScale: Exit Sub

XMin = CDbl(txtXMinimum)
XMax = CDbl(txtXMaximum)
YMin = CDbl(txtYMinimum)
YMax = CDbl(txtYMaximum)
XScl = CDbl(txtXScale)
YScl = CDbl(txtYScale)

Unload Me
End Sub

Private Sub Form_Load()
txtXMinimum = CStr(XMin)
txtXMaximum = CStr(XMax)
txtYMinimum = CStr(YMin)
txtYMaximum = CStr(YMax)
txtXScale = CStr(XScl)
txtYScale = CStr(YScl)
End Sub

Private Sub txtXMaximum_LostFocus()
CheckTBoxes
End Sub

Private Sub txtXMinimum_LostFocus()
CheckTBoxes
End Sub

Private Sub txtXScale_LostFocus()
CheckTBoxes
End Sub

Private Sub txtYMaximum_LostFocus()
CheckTBoxes
End Sub

Private Sub txtYMinimum_LostFocus()
CheckTBoxes
End Sub

Private Sub txtYScale_LostFocus()
CheckTBoxes
End Sub

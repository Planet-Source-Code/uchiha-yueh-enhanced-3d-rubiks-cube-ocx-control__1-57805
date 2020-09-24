VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rubik's Cube"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5595
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   5595
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFCWise 
      Caption         =   "<<"
      Height          =   255
      Left            =   1320
      TabIndex        =   21
      ToolTipText     =   "Clock-Wise Rotation"
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton cmdFCCWise 
      Caption         =   ">>"
      Height          =   255
      Left            =   3480
      TabIndex        =   20
      ToolTipText     =   "Counter Clock-Wise Rotation"
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton cmdBCCWise 
      Caption         =   ">>"
      Height          =   255
      Left            =   3480
      TabIndex        =   19
      ToolTipText     =   "Counter Clock-Wise Rotation"
      Top             =   3735
      Width           =   855
   End
   Begin VB.CommandButton cmdBCWise 
      Caption         =   "<<"
      Height          =   255
      Left            =   1320
      TabIndex        =   18
      ToolTipText     =   "Clock-Wise Rotation"
      Top             =   3735
      Width           =   855
   End
   Begin VB.CommandButton cmdTopCWise 
      Caption         =   "<<"
      Height          =   255
      Left            =   1320
      TabIndex        =   15
      ToolTipText     =   "Clock-Wise Rotation"
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdTopCCWise 
      Caption         =   ">>"
      Height          =   255
      Left            =   3480
      TabIndex        =   14
      ToolTipText     =   "Counter Clock-Wise Rotation"
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdBottomCCWise 
      Caption         =   ">>"
      Height          =   255
      Left            =   3480
      TabIndex        =   13
      ToolTipText     =   "Counter Clock-Wise Rotation"
      Top             =   375
      Width           =   855
   End
   Begin VB.CommandButton cmdBottomCWise 
      Caption         =   "<<"
      Height          =   255
      Left            =   1320
      TabIndex        =   12
      ToolTipText     =   "Clock-Wise Rotation"
      Top             =   375
      Width           =   855
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3960
      Top             =   4080
   End
   Begin VB.CommandButton cmdRCWise 
      Caption         =   "<<"
      Height          =   255
      Left            =   4560
      TabIndex        =   10
      ToolTipText     =   "Clock-Wise Rotation"
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton cmdRCCWise 
      Caption         =   ">>"
      Height          =   255
      Left            =   4560
      TabIndex        =   9
      ToolTipText     =   "Counter Clock-Wise Rotation"
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton cmdLCWise 
      Caption         =   "<<"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      ToolTipText     =   "Clock-Wise Rotation"
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton cmdLCCWise 
      Caption         =   ">>"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      ToolTipText     =   "Counter Clock-Wise Rotation"
      Top             =   2280
      Width           =   855
   End
   Begin RubiksCube.RubikCube RubikCube1 
      Height          =   2655
      Left            =   1320
      TabIndex        =   2
      Top             =   720
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4683
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4560
      Top             =   2760
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Reset"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Scramble"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Top"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   2280
      TabIndex        =   17
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bottom"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   2280
      TabIndex        =   16
      Top             =   375
      Width           =   1095
   End
   Begin VB.Label lblTimer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time elapsed: 00:00:00"
      Height          =   195
      Left            =   1920
      TabIndex        =   11
      Top             =   4200
      Width           =   1665
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Right"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   4440
      TabIndex        =   6
      Top             =   1980
      Width           =   1095
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Left"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   1980
      Width           =   1095
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Back"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   4
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Front"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   2280
      TabIndex        =   3
      Top             =   3480
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hr As Integer, min As Integer, sec As Integer

Private Sub cmdBCCWise_Click()

Dim Z As String

'call EnableControl procedure..
Call EnableControl

'disable form..
Me.Enabled = False

'enabled Timer1 control..
Timer1.Enabled = True

X = RubikCube1.GetXCoor
Y = RubikCube1.GetYCoor
Z = X & "/" & Y
RubikCube1.GetRotation Z, "BA", "CCWise"

End Sub

Private Sub cmdBCWise_Click()

Dim Z As String

'call EnableControl procedure..
Call EnableControl

'disable form..
Me.Enabled = False

'enabled Timer1 control..
Timer1.Enabled = True

X = RubikCube1.GetXCoor
Y = RubikCube1.GetYCoor
Z = X & "/" & Y
RubikCube1.GetRotation Z, "BA", "CWise"

End Sub

Private Sub cmdBottomCCWise_Click()
Dim Z As String

'call EnableControl procedure..
Call EnableControl

'disable form..
Me.Enabled = False

'enabled Timer1 control..
Timer1.Enabled = True

X = RubikCube1.GetXCoor
Y = RubikCube1.GetYCoor
Z = X & "/" & Y
RubikCube1.GetRotation Z, "BO", "CCWise"

End Sub

Private Sub cmdBottomCWise_Click()

Dim Z As String

'call EnableControl procedure..
Call EnableControl

'disable form..
Me.Enabled = False

'enabled Timer1 control..
Timer1.Enabled = True

X = RubikCube1.GetXCoor
Y = RubikCube1.GetYCoor
Z = X & "/" & Y
RubikCube1.GetRotation Z, "BO", "CWise"

End Sub

Private Sub cmdFCCWise_Click()

Dim Z As String

'call EnableControl procedure..
Call EnableControl

'disable form..
Me.Enabled = False

'enabled Timer1 control..
Timer1.Enabled = True

X = RubikCube1.GetXCoor
Y = RubikCube1.GetYCoor
Z = X & "/" & Y
RubikCube1.GetRotation Z, "F", "CCWise"

End Sub

Private Sub cmdFCWise_Click()

Dim Z As String

'call EnableControl procedure..
Call EnableControl

'disable form..
Me.Enabled = False

'enabled Timer1 control..
Timer1.Enabled = True

X = RubikCube1.GetXCoor
Y = RubikCube1.GetYCoor
Z = X & "/" & Y
RubikCube1.GetRotation Z, "F", "CWise"

End Sub

Private Sub cmdLCCWise_Click()
Dim Z As String

'call EnableControl procedure..
Call EnableControl

'disable form..
Me.Enabled = False

'enabled Timer1 control..
Timer1.Enabled = True

X = RubikCube1.GetXCoor
Y = RubikCube1.GetYCoor
Z = X & "/" & Y
RubikCube1.GetRotation Z, "L", "CCWise"

End Sub

Private Sub cmdLCWise_Click()
Dim Z As String

'call EnableControl procedure..
Call EnableControl

'disable form..
Me.Enabled = False

'enabled Timer1 control..
Timer1.Enabled = True

X = RubikCube1.GetXCoor
Y = RubikCube1.GetYCoor
Z = X & "/" & Y
RubikCube1.GetRotation Z, "L", "CWise"

End Sub

Private Sub cmdRCCWise_Click()
Dim Z As String

'call EnableControl procedure..
Call EnableControl

'disable form..
Me.Enabled = False

'enabled Timer1 control..
Timer1.Enabled = True

X = RubikCube1.GetXCoor
Y = RubikCube1.GetYCoor
Z = X & "/" & Y
RubikCube1.GetRotation Z, "R", "CCWise"

End Sub

Private Sub cmdRCWise_Click()

Dim Z As String

'call EnableControl procedure..
Call EnableControl

'disable form..
Me.Enabled = False

'enabled Timer1 control..
Timer1.Enabled = True

X = RubikCube1.GetXCoor
Y = RubikCube1.GetYCoor
Z = X & "/" & Y
RubikCube1.GetRotation Z, "R", "CWise"

End Sub

Private Sub cmdTopCCWise_Click()

Dim Z As String

'call EnableControl procedure..
Call EnableControl

'disable form..
Me.Enabled = False

'enabled Timer1 control..
Timer1.Enabled = True

X = RubikCube1.GetXCoor
Y = RubikCube1.GetYCoor
Z = X & "/" & Y
RubikCube1.GetRotation Z, "T", "CCWise"

End Sub

Private Sub cmdTopCWise_Click()
Dim Z As String

'call EnableControl procedure..
Call EnableControl

'disable form..
Me.Enabled = False

'enabled Timer1 control..
Timer1.Enabled = True

X = RubikCube1.GetXCoor
Y = RubikCube1.GetYCoor
Z = X & "/" & Y
RubikCube1.GetRotation Z, "T", "CWise"

End Sub

Private Sub Command1_Click()

'if Time Elapsed is enabled..
If Timer2.Enabled = True Then
    
    'display elapsed time..
    MsgBox lblTimer.Caption, , Me.Caption

    'ask if user want to scramble cube again..
    ans = MsgBox("Are you sure?", _
        vbQuestion + vbYesNo, "Scramble Cube")

    'if user answered "NO", exit to this event..
    If ans = vbNo Then Exit Sub

End If

'scramble cube..
RubikCube1.ScrambleCube

'display that timer has been started..
MsgBox "Timer starts now!!", vbInformation, Me.Caption

'reset all values..
hr = 0: min = 0: sec = 0

'reset lblTimer to default caption..
lblTimer.Caption = "Time elapsed: 00:00:00"

'enable Timer2 control..
Timer2.Enabled = True

End Sub

Private Sub Command2_Click()

'ask if Time Elapsed is enabled..
If Timer2.Enabled = True Then
    
    'ask if user wants to reset cube..
    ans = MsgBox("Are you sure?", _
        vbQuestion + vbYesNo, "Reset Cube")

    'if user answered "NO", exit to this event..
    If ans = vbNo Then Exit Sub
    
    'display elpased time..
    MsgBox lblTimer.Caption, , Me.Caption
    
    'disable Timer2..
    Timer2.Enabled = False

End If

'reset cube..
RubikCube1.ResetCube

'reset lblTimer to default caption..
lblTimer.Caption = "Time elapsed: 00:00:00"

'display message..
MsgBox "Rubik's cube has been reset.", vbInformation, Me.Caption

End Sub

Private Sub Form_Load()

hr = 0
min = 0
sec = 0

End Sub

Private Sub Timer1_Timer()

Me.Enabled = True
Timer1.Enabled = False

Call EnableControl

End Sub

Sub EnableControl()

Dim Ctrl As Control

'check each control in a form..
For Each Ctrl In Controls
    
    'ask if the type of control is a command button..
    If TypeOf Ctrl Is CommandButton Then
    
        'the NOT operator inverts the value
        'of the ctrl. If ctrl is enabled, disabled it
        'and vice versa..
        
        Ctrl.Enabled = Not Ctrl.Enabled
    
    End If

Next

End Sub

Private Sub Timer2_Timer()
Dim h As String, m As String, s As String

'initialize variables..
h = "": m = "": s = ""

'add 1 to seconds..
sec = sec + 1


If sec = 60 Then
    
    sec = 0
    min = min + 1
    
    If min = 60 Then
    
        min = 0
        hr = hr + 1
    
    End If
    
End If

If sec < 10 Then
    s = "0" & LTrim(Str(sec))
Else
    s = LTrim(Str(sec))
End If

If min < 10 Then
    m = "0" & LTrim(Str(min))
Else
    m = LTrim(Str(min))
End If

If hr < 10 Then
    h = "0" & LTrim(Str(hr))
Else
    h = LTrim(Str(hr))
End If

lblTimer.Caption = _
    "Time elapsed: " & h & ":" & m & ":" & s

End Sub

Sub Check_Answer()

If RubikCube1.GetCube = "RRRRRRRRRYYYYYYYYYPPPPPPPPPWWWWWWWWWBBBBBBBBBGGGGGGGGG" Then
    Timer2.Enabled = False
    MsgBox "You have complete Rubik's Cube!!" & _
        vbCrLf & lblTimer.Caption, vbInformation, "Congratulations"

    'reset lblTimer to default caption..
    lblTimer.Caption = "Time elapsed: 00:00:00"

End If

End Sub


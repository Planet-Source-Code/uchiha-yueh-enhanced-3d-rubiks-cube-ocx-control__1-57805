VERSION 5.00
Begin VB.UserControl RubikCube 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   ClientHeight    =   3645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3165
   LockControls    =   -1  'True
   ScaleHeight     =   3645
   ScaleWidth      =   3165
   ToolboxBitmap   =   "RubikCube.ctx":0000
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   5040
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.PictureBox Picture2 
      Enabled         =   0   'False
      Height          =   1695
      Left            =   -3600
      ScaleHeight     =   1635
      ScaleWidth      =   2715
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   2775
      Begin VB.HScrollBar Rott1 
         Height          =   240
         LargeChange     =   45
         Left            =   0
         Max             =   180
         Min             =   -180
         SmallChange     =   15
         TabIndex        =   10
         Top             =   0
         Value           =   -30
         Width           =   2880
      End
      Begin VB.VScrollBar Rott2 
         Height          =   2025
         LargeChange     =   45
         Left            =   120
         Max             =   180
         Min             =   -180
         SmallChange     =   15
         TabIndex        =   9
         Top             =   240
         Value           =   30
         Width           =   240
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      Height          =   4185
      Left            =   -1560
      ScaleHeight     =   279
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   296
      TabIndex        =   0
      Top             =   -1560
      Width           =   4440
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   0
         TabIndex        =   7
         Text            =   "99"
         Top             =   -360
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.ListBox List3 
         Height          =   1620
         Left            =   -1080
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   0
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton Command26 
         Enabled         =   0   'False
         Height          =   345
         Left            =   3600
         TabIndex        =   1
         Top             =   4920
         Width           =   360
      End
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   1200
         Top             =   -480
      End
      Begin VB.Label lblRotate 
         BackStyle       =   0  'Transparent
         Height          =   375
         Index           =   0
         Left            =   1920
         MouseIcon       =   "RubikCube.ctx":0312
         MousePointer    =   99  'Custom
         TabIndex        =   5
         ToolTipText     =   "Rotate cube upward.."
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label lblRotate 
         BackStyle       =   0  'Transparent
         Height          =   375
         Index           =   1
         Left            =   1920
         MouseIcon       =   "RubikCube.ctx":0754
         MousePointer    =   99  'Custom
         TabIndex        =   4
         ToolTipText     =   "Rotate cube downward.."
         Top             =   3720
         Width           =   2055
      End
      Begin VB.Label lblRotate 
         BackStyle       =   0  'Transparent
         Height          =   1575
         Index           =   2
         Left            =   3840
         MouseIcon       =   "RubikCube.ctx":0B96
         MousePointer    =   99  'Custom
         TabIndex        =   3
         ToolTipText     =   "Rotate cube to the right.."
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label lblRotate 
         BackStyle       =   0  'Transparent
         Height          =   1575
         Index           =   3
         Left            =   1440
         MouseIcon       =   "RubikCube.ctx":0FD8
         MousePointer    =   99  'Custom
         TabIndex        =   2
         ToolTipText     =   "Rotate cube to the left.."
         Top             =   2040
         Width           =   495
      End
   End
   Begin VB.Label lblY 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   12
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   11
      Top             =   2880
      Width           =   1455
   End
End
Attribute VB_Name = "RubikCube"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'3D Drawing
'================================================
'Frames Structure
Private Type Frames
    X1 As Single
    Y1 As Single
    Z1 As Single
    X2 As Single
    Y2 As Single
    Z2 As Single
    X3 As Single
    Y3 As Single
    Z3 As Single
    X4 As Single
    Y4 As Single
    Z4 As Single
End Type
'================================================

'================================================
'Pnt Structure
Private Type Pnt
    X As Single
    Y As Single
    Z As Single
End Type
'================================================


'================================================
'Variables..
Private CubeA As String
Private SolPos As Long
Private RotateVal As Integer
Private RotateSgn As String

Private ori(999) As Frames
Private dup(999) As Frames
Private colo(999) As Long
Private Pont(8) As Pnt
Private XY As String
Private Movement As String
'================================================

Private Sub Init3D()

For i = 1 To 9
    
    ori(i).Z1 = -150
    ori(i).Z2 = -150
    ori(i).Z3 = -150
    ori(i).Z4 = -150
    ori(i).X1 = -150 + ((i - 1) Mod 3) * 100
    ori(i).X2 = -50 + ((i - 1) Mod 3) * 100
    ori(i).X3 = -50 + ((i - 1) Mod 3) * 100
    ori(i).X4 = -150 + ((i - 1) Mod 3) * 100
    ori(i).Y1 = -150 + ((i - 1) \ 3) * 100
    ori(i).Y2 = -150 + ((i - 1) \ 3) * 100
    ori(i).Y3 = -50 + ((i - 1) \ 3) * 100
    ori(i).Y4 = -50 + ((i - 1) \ 3) * 100

Next

For i = 10 To 18
    
    ori(i).X1 = 150
    ori(i).X2 = 150
    ori(i).X3 = 150
    ori(i).X4 = 150
    ori(i).Z1 = -150 + ((i - 1) Mod 3) * 100
    ori(i).Z2 = -50 + ((i - 1) Mod 3) * 100
    ori(i).Z3 = -50 + ((i - 1) Mod 3) * 100
    ori(i).Z4 = -150 + ((i - 1) Mod 3) * 100
    ori(i).Y1 = -150 + ((i - 10) \ 3) * 100
    ori(i).Y2 = -150 + ((i - 10) \ 3) * 100
    ori(i).Y3 = -50 + ((i - 10) \ 3) * 100
    ori(i).Y4 = -50 + ((i - 10) \ 3) * 100

Next

For i = 19 To 27

    ori(i).Z1 = 150
    ori(i).Z2 = 150
    ori(i).Z3 = 150
    ori(i).Z4 = 150
    ori(i).X1 = 150 - ((i - 1) Mod 3) * 100
    ori(i).X2 = 50 - ((i - 1) Mod 3) * 100
    ori(i).X3 = 50 - ((i - 1) Mod 3) * 100
    ori(i).X4 = 150 - ((i - 1) Mod 3) * 100
    ori(i).Y1 = -150 + ((i - 19) \ 3) * 100
    ori(i).Y2 = -150 + ((i - 19) \ 3) * 100
    ori(i).Y3 = -50 + ((i - 19) \ 3) * 100
    ori(i).Y4 = -50 + ((i - 19) \ 3) * 100

Next


For i = 28 To 36

    ori(i).X1 = -150
    ori(i).X2 = -150
    ori(i).X3 = -150
    ori(i).X4 = -150
    ori(i).Z1 = 150 - ((i - 1) Mod 3) * 100
    ori(i).Z2 = 50 - ((i - 1) Mod 3) * 100
    ori(i).Z3 = 50 - ((i - 1) Mod 3) * 100
    ori(i).Z4 = 150 - ((i - 1) Mod 3) * 100
    ori(i).Y1 = -150 + ((i - 28) \ 3) * 100
    ori(i).Y2 = -150 + ((i - 28) \ 3) * 100
    ori(i).Y3 = -50 + ((i - 28) \ 3) * 100
    ori(i).Y4 = -50 + ((i - 28) \ 3) * 100

Next

For i = 37 To 45
    
    ori(i).Y1 = -150
    ori(i).Y2 = -150
    ori(i).Y3 = -150
    ori(i).Y4 = -150
    ori(i).X1 = -150 + ((i - 1) Mod 3) * 100
    ori(i).X2 = -50 + ((i - 1) Mod 3) * 100
    ori(i).X3 = -50 + ((i - 1) Mod 3) * 100
    ori(i).X4 = -150 + ((i - 1) Mod 3) * 100
    ori(i).Z1 = 150 - ((i - 37) \ 3) * 100
    ori(i).Z2 = 150 - ((i - 37) \ 3) * 100
    ori(i).Z3 = 50 - ((i - 37) \ 3) * 100
    ori(i).Z4 = 50 - ((i - 37) \ 3) * 100
Next

For i = 46 To 54
    
    ori(i).Y1 = 150
    ori(i).Y2 = 150
    ori(i).Y3 = 150
    ori(i).Y4 = 150
    ori(i).X1 = -150 + ((i - 1) Mod 3) * 100
    ori(i).X2 = -50 + ((i - 1) Mod 3) * 100
    ori(i).X3 = -50 + ((i - 1) Mod 3) * 100
    ori(i).X4 = -150 + ((i - 1) Mod 3) * 100
    ori(i).Z1 = 150 - ((i - 46) \ 3) * 100
    ori(i).Z2 = 150 - ((i - 46) \ 3) * 100
    ori(i).Z3 = 50 - ((i - 46) \ 3) * 100
    ori(i).Z4 = 50 - ((i - 46) \ 3) * 100

Next

For i = 1 To Len(CubeA)
    
    Select Case Mid(CubeA, i, 1)
        
        Case "R"
            colo(i) = RGB(255, 0, 0)
        Case "Y"
            colo(i) = RGB(255, 255, 0)
        Case "P"
            colo(i) = &H80FF&
        Case "W"
            colo(i) = RGB(255, 255, 255)
        Case "B"
            colo(i) = RGB(50, 50, 200)
        Case "G"
            colo(i) = RGB(50, 200, 50)

    End Select

Next

End Sub

Public Sub RotateFrontCWise()

Dim Fra(21) As Integer

Fra(1) = 1
Fra(2) = 2
Fra(3) = 3
Fra(4) = 4
Fra(5) = 5
Fra(6) = 6
Fra(7) = 7
Fra(8) = 8
Fra(9) = 9
Fra(10) = 10
Fra(11) = 13
Fra(12) = 16
Fra(13) = 43
Fra(14) = 44
Fra(15) = 45
Fra(16) = 30
Fra(17) = 33
Fra(18) = 36
Fra(19) = 52
Fra(20) = 53
Fra(21) = 54

For ii = 1 To 4

    Call Init3D
    
    For j = 1 To 21
        Rotate ori(Fra(j)).X1, ori(Fra(j)).Y1, 3.14159265258979 * 22.5 * ii / 180
        Rotate ori(Fra(j)).X2, ori(Fra(j)).Y2, 3.14159265258979 * 22.5 * ii / 180
        Rotate ori(Fra(j)).X3, ori(Fra(j)).Y3, 3.14159265258979 * 22.5 * ii / 180
        Rotate ori(Fra(j)).X4, ori(Fra(j)).Y4, 3.14159265258979 * 22.5 * ii / 180
    Next
    
    For i = 1 To 54
        Rotate ori(i).Z1, ori(i).X1, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z2, ori(i).X2, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z3, ori(i).X3, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z4, ori(i).X4, 3.14159265258979 * -Rott1.Value / 180
    Next
    
    For i = 1 To 54
        Rotate ori(i).Y1, ori(i).Z1, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y2, ori(i).Z2, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y3, ori(i).Z3, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y4, ori(i).Z4, 3.14159265258979 * Rott2.Value / 180
    Next
    
    DoEvents
    Picture1.Cls
    Picture1.Scale (-700, -700)-(300, 300)
    
    Call DrawCube

Next

Call RotateFace(CubeA, "f")
Call ViewCube(CubeA)
Call Rott1_Scroll

End Sub

Private Sub ViewCube(Cube As String)

Cube = UCase(Cube)
Call Rott1_Scroll

End Sub
Private Sub DrawCube()

For i = 1 To 54
    List3.AddItem Str(FrDepth(ori(i))) & " " & Str(i)
Next

For i = 1 To 54
    j = Val(Right(List3.List(i), 3))
    DrawPolygon ori(j), colo(j)
Next

DoEvents
List3.Clear

End Sub

Private Sub Rott1_Change()

If Rott1.Value = 90 Or _
    Rott1.Value = -90 Or _
    Rott1.Value = 135 Or _
    Rott1.Value = -135 Then
        If RotateSgn = "+" Then
            Rott1.Value = Rott1.Value + 30
        Else
            Rott1.Value = Rott1.Value - 30
        End If
End If

Call Init3D

For i = 1 To 54
    Rotate ori(i).Z1, ori(i).X1, 3.14159265258979 * -Rott1.Value / 180
    Rotate ori(i).Z2, ori(i).X2, 3.14159265258979 * -Rott1.Value / 180
    Rotate ori(i).Z3, ori(i).X3, 3.14159265258979 * -Rott1.Value / 180
    Rotate ori(i).Z4, ori(i).X4, 3.14159265258979 * -Rott1.Value / 180
Next

For i = 1 To 54
    Rotate ori(i).Y1, ori(i).Z1, 3.14159265258979 * Rott2.Value / 180
    Rotate ori(i).Y2, ori(i).Z2, 3.14159265258979 * Rott2.Value / 180
    Rotate ori(i).Y3, ori(i).Z3, 3.14159265258979 * Rott2.Value / 180
    Rotate ori(i).Y4, ori(i).Z4, 3.14159265258979 * Rott2.Value / 180
Next

DoEvents
Picture1.Cls
Picture1.Scale (-700, -700)-(300, 300)

Call DrawCube
Picture1.Scale (-300, -300)-(700, 700)

lblX.Caption = "X=" & Rott1.Value

End Sub

Private Sub Rott1_Scroll()

Call Rott1_Change

End Sub

Private Sub Rott2_Change()

If Rott2.Value = 90 Or _
    Rott2.Value = -90 Or _
    Rott2.Value = 0 Then
        If RotateSgn = "+" Then
            Rott2.Value = Rott2.Value + 30
        Else
            Rott2.Value = Rott2.Value - 30
        End If
End If

Call Init3D

For i = 1 To 54
    Rotate ori(i).Z1, ori(i).X1, 3.14159265258979 * -Rott1.Value / 180
    Rotate ori(i).Z2, ori(i).X2, 3.14159265258979 * -Rott1.Value / 180
    Rotate ori(i).Z3, ori(i).X3, 3.14159265258979 * -Rott1.Value / 180
    Rotate ori(i).Z4, ori(i).X4, 3.14159265258979 * -Rott1.Value / 180
Next

For i = 1 To 54
    Rotate ori(i).Y1, ori(i).Z1, 3.14159265258979 * Rott2.Value / 180
    Rotate ori(i).Y2, ori(i).Z2, 3.14159265258979 * Rott2.Value / 180
    Rotate ori(i).Y3, ori(i).Z3, 3.14159265258979 * Rott2.Value / 180
    Rotate ori(i).Y4, ori(i).Z4, 3.14159265258979 * Rott2.Value / 180
Next

DoEvents
Picture1.Cls
Picture1.Scale (-700, -700)-(300, 300)

Call DrawCube
Picture1.Scale (-300, -300)-(700, 700)

lblY.Caption = "Y=" & Rott2.Value

End Sub

Private Sub Rott2_Scroll()

Call Rott2_Change

End Sub

Private Function FrDepth(Fram As Frames) As Double

    xx = (Fram.X1 + Fram.X2 + Fram.X3 + Fram.X4) / 4
    yy = (Fram.Y1 + Fram.Y2 + Fram.Y3 + Fram.Y4) / 4
    zz = (Fram.Z1 + Fram.Z2 + Fram.Z3 + Fram.Z4) / 4
    FrDepth = xx ^ 2 + yy ^ 2 + (zz - 600) ^ 2

End Function

Public Sub RotateFrontCCWise()

Dim Fra(21) As Integer

Fra(1) = 1
Fra(2) = 2
Fra(3) = 3
Fra(4) = 4
Fra(5) = 5
Fra(6) = 6
Fra(7) = 7
Fra(8) = 8
Fra(9) = 9
Fra(10) = 10
Fra(11) = 13
Fra(12) = 16
Fra(13) = 43
Fra(14) = 44
Fra(15) = 45
Fra(16) = 30
Fra(17) = 33
Fra(18) = 36
Fra(19) = 52
Fra(20) = 53
Fra(21) = 54

For ii = 1 To 4
    
    Call Init3D
    
    For j = 1 To 21
        Rotate ori(Fra(j)).X1, ori(Fra(j)).Y1, 3.14159265258979 * -22.5 * ii / 180
        Rotate ori(Fra(j)).X2, ori(Fra(j)).Y2, 3.14159265258979 * -22.5 * ii / 180
        Rotate ori(Fra(j)).X3, ori(Fra(j)).Y3, 3.14159265258979 * -22.5 * ii / 180
        Rotate ori(Fra(j)).X4, ori(Fra(j)).Y4, 3.14159265258979 * -22.5 * ii / 180
    Next
    
    For i = 1 To 54
        Rotate ori(i).Z1, ori(i).X1, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z2, ori(i).X2, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z3, ori(i).X3, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z4, ori(i).X4, 3.14159265258979 * -Rott1.Value / 180
    Next
    
    For i = 1 To 54
        Rotate ori(i).Y1, ori(i).Z1, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y2, ori(i).Z2, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y3, ori(i).Z3, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y4, ori(i).Z4, 3.14159265258979 * Rott2.Value / 180
    Next
    
    DoEvents
    Picture1.Cls
    Picture1.Scale (-700, -700)-(300, 300)
    Call DrawCube

Next


Call RotateFace(CubeA, "f")
Call RotateFace(CubeA, "f")
Call RotateFace(CubeA, "f")
Call ViewCube(CubeA)
Call Rott1_Scroll

End Sub

Public Sub RotateBackCWise()

Dim Fra(21) As Integer

Fra(1) = 19
Fra(2) = 20
Fra(3) = 21
Fra(4) = 22
Fra(5) = 23
Fra(6) = 24
Fra(7) = 25
Fra(8) = 26
Fra(9) = 27
Fra(10) = 12
Fra(11) = 15
Fra(12) = 18
Fra(13) = 37
Fra(14) = 38
Fra(15) = 39
Fra(16) = 28
Fra(17) = 31
Fra(18) = 34
Fra(19) = 46
Fra(20) = 47
Fra(21) = 48

For ii = 1 To 4

    Call Init3D
    
    For j = 1 To 21
        Rotate ori(Fra(j)).Y1, ori(Fra(j)).X1, 3.14159265258979 * 22.5 * ii / 180
        Rotate ori(Fra(j)).Y2, ori(Fra(j)).X2, 3.14159265258979 * 22.5 * ii / 180
        Rotate ori(Fra(j)).Y3, ori(Fra(j)).X3, 3.14159265258979 * 22.5 * ii / 180
        Rotate ori(Fra(j)).Y4, ori(Fra(j)).X4, 3.14159265258979 * 22.5 * ii / 180
    Next
    
    For i = 1 To 54
        Rotate ori(i).Z1, ori(i).X1, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z2, ori(i).X2, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z3, ori(i).X3, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z4, ori(i).X4, 3.14159265258979 * -Rott1.Value / 180
    Next
    
    For i = 1 To 54
        Rotate ori(i).Y1, ori(i).Z1, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y2, ori(i).Z2, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y3, ori(i).Z3, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y4, ori(i).Z4, 3.14159265258979 * Rott2.Value / 180
    Next
    
    DoEvents
    Picture1.Cls
    Picture1.Scale (-700, -700)-(300, 300)
    
    Call DrawCube

Next


Call RotateFace(CubeA, "b")
Call ViewCube(CubeA)

End Sub

Public Sub RotateBackCCWise()

Dim Fra(21) As Integer

Fra(1) = 19
Fra(2) = 20
Fra(3) = 21
Fra(4) = 22
Fra(5) = 23
Fra(6) = 24
Fra(7) = 25
Fra(8) = 26
Fra(9) = 27
Fra(10) = 12
Fra(11) = 15
Fra(12) = 18
Fra(13) = 37
Fra(14) = 38
Fra(15) = 39
Fra(16) = 28
Fra(17) = 31
Fra(18) = 34
Fra(19) = 46
Fra(20) = 47
Fra(21) = 48

For ii = 1 To 4
    
    Call Init3D
    
    For j = 1 To 21
        Rotate ori(Fra(j)).Y1, ori(Fra(j)).X1, 3.14159265258979 * -22.5 * ii / 180
        Rotate ori(Fra(j)).Y2, ori(Fra(j)).X2, 3.14159265258979 * -22.5 * ii / 180
        Rotate ori(Fra(j)).Y3, ori(Fra(j)).X3, 3.14159265258979 * -22.5 * ii / 180
        Rotate ori(Fra(j)).Y4, ori(Fra(j)).X4, 3.14159265258979 * -22.5 * ii / 180
    Next
    For i = 1 To 54
        Rotate ori(i).Z1, ori(i).X1, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z2, ori(i).X2, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z3, ori(i).X3, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z4, ori(i).X4, 3.14159265258979 * -Rott1.Value / 180
    Next
    For i = 1 To 54
        Rotate ori(i).Y1, ori(i).Z1, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y2, ori(i).Z2, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y3, ori(i).Z3, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y4, ori(i).Z4, 3.14159265258979 * Rott2.Value / 180
    Next
    
    DoEvents
    Picture1.Cls
    Picture1.Scale (-700, -700)-(300, 300)
    Call DrawCube

Next

Call RotateFace(CubeA, "b")
Call RotateFace(CubeA, "b")
Call RotateFace(CubeA, "b")
Call ViewCube(CubeA)

End Sub

Public Sub RotateTopCWise()

Dim Fra(21) As Integer

Fra(1) = 37
Fra(2) = 38
Fra(3) = 39
Fra(4) = 40
Fra(5) = 41
Fra(6) = 42
Fra(7) = 43
Fra(8) = 44
Fra(9) = 45
Fra(10) = 1
Fra(11) = 2
Fra(12) = 3
Fra(13) = 10
Fra(14) = 11
Fra(15) = 12
Fra(16) = 19
Fra(17) = 20
Fra(18) = 21
Fra(19) = 28
Fra(20) = 29
Fra(21) = 30

For ii = 1 To 4
    
    Call Init3D
    
    For j = 1 To 21
        Rotate ori(Fra(j)).Z1, ori(Fra(j)).X1, 3.14159265258979 * 22.5 * ii / 180
        Rotate ori(Fra(j)).Z2, ori(Fra(j)).X2, 3.14159265258979 * 22.5 * ii / 180
        Rotate ori(Fra(j)).Z3, ori(Fra(j)).X3, 3.14159265258979 * 22.5 * ii / 180
        Rotate ori(Fra(j)).Z4, ori(Fra(j)).X4, 3.14159265258979 * 22.5 * ii / 180
    Next
    For i = 1 To 54
        Rotate ori(i).Z1, ori(i).X1, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z2, ori(i).X2, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z3, ori(i).X3, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z4, ori(i).X4, 3.14159265258979 * -Rott1.Value / 180
    Next
    For i = 1 To 54
        Rotate ori(i).Y1, ori(i).Z1, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y2, ori(i).Z2, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y3, ori(i).Z3, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y4, ori(i).Z4, 3.14159265258979 * Rott2.Value / 180
    Next
    
    DoEvents
    Picture1.Cls
    Picture1.Scale (-700, -700)-(300, 300)
    Call DrawCube

Next

Call RotateFace(CubeA, "t")
Call ViewCube(CubeA)

End Sub

Public Sub RotateTopCCWise()

Dim Fra(21) As Integer

Fra(1) = 37
Fra(2) = 38
Fra(3) = 39
Fra(4) = 40
Fra(5) = 41
Fra(6) = 42
Fra(7) = 43
Fra(8) = 44
Fra(9) = 45
Fra(10) = 1
Fra(11) = 2
Fra(12) = 3
Fra(13) = 10
Fra(14) = 11
Fra(15) = 12
Fra(16) = 19
Fra(17) = 20
Fra(18) = 21
Fra(19) = 28
Fra(20) = 29
Fra(21) = 30

For ii = 1 To 4

    Call Init3D
    For j = 1 To 21
        Rotate ori(Fra(j)).Z1, ori(Fra(j)).X1, 3.14159265258979 * -22.5 * ii / 180
        Rotate ori(Fra(j)).Z2, ori(Fra(j)).X2, 3.14159265258979 * -22.5 * ii / 180
        Rotate ori(Fra(j)).Z3, ori(Fra(j)).X3, 3.14159265258979 * -22.5 * ii / 180
        Rotate ori(Fra(j)).Z4, ori(Fra(j)).X4, 3.14159265258979 * -22.5 * ii / 180
    Next
    For i = 1 To 54
        Rotate ori(i).Z1, ori(i).X1, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z2, ori(i).X2, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z3, ori(i).X3, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z4, ori(i).X4, 3.14159265258979 * -Rott1.Value / 180
    Next
    For i = 1 To 54
        Rotate ori(i).Y1, ori(i).Z1, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y2, ori(i).Z2, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y3, ori(i).Z3, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y4, ori(i).Z4, 3.14159265258979 * Rott2.Value / 180
    Next
    
    DoEvents
    Picture1.Cls
    Picture1.Scale (-700, -700)-(300, 300)
    Call DrawCube

Next

Call RotateFace(CubeA, "t")
Call RotateFace(CubeA, "t")
Call RotateFace(CubeA, "t")
Call ViewCube(CubeA)
End Sub

Public Sub RotateBottomCWise()

Dim Fra(21) As Integer

Fra(1) = 46
Fra(2) = 47
Fra(3) = 48
Fra(4) = 49
Fra(5) = 50
Fra(6) = 51
Fra(7) = 52
Fra(8) = 53
Fra(9) = 54
Fra(10) = 7
Fra(11) = 8
Fra(12) = 9
Fra(13) = 16
Fra(14) = 17
Fra(15) = 18
Fra(16) = 25
Fra(17) = 26
Fra(18) = 27
Fra(19) = 34
Fra(20) = 35
Fra(21) = 36

For ii = 1 To 4

    Call Init3D
    
    For j = 1 To 21
        Rotate ori(Fra(j)).Z1, ori(Fra(j)).X1, 3.14159265258979 * -22.5 * ii / 180
        Rotate ori(Fra(j)).Z2, ori(Fra(j)).X2, 3.14159265258979 * -22.5 * ii / 180
        Rotate ori(Fra(j)).Z3, ori(Fra(j)).X3, 3.14159265258979 * -22.5 * ii / 180
        Rotate ori(Fra(j)).Z4, ori(Fra(j)).X4, 3.14159265258979 * -22.5 * ii / 180
    Next
    
    For i = 1 To 54
        Rotate ori(i).Z1, ori(i).X1, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z2, ori(i).X2, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z3, ori(i).X3, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z4, ori(i).X4, 3.14159265258979 * -Rott1.Value / 180
    Next
    
    For i = 1 To 54
        Rotate ori(i).Y1, ori(i).Z1, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y2, ori(i).Z2, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y3, ori(i).Z3, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y4, ori(i).Z4, 3.14159265258979 * Rott2.Value / 180
    Next
    
    DoEvents
    Picture1.Cls
    Picture1.Scale (-700, -700)-(300, 300)
    
    Call DrawCube

Next

Call RotateFace(CubeA, "d")
Call ViewCube(CubeA)

End Sub

Public Sub RotateBottomCCWise()

Dim Fra(21) As Integer

Fra(1) = 46
Fra(2) = 47
Fra(3) = 48
Fra(4) = 49
Fra(5) = 50
Fra(6) = 51
Fra(7) = 52
Fra(8) = 53
Fra(9) = 54
Fra(10) = 7
Fra(11) = 8
Fra(12) = 9
Fra(13) = 16
Fra(14) = 17
Fra(15) = 18
Fra(16) = 25
Fra(17) = 26
Fra(18) = 27
Fra(19) = 34
Fra(20) = 35
Fra(21) = 36

For ii = 1 To 4
    
    Call Init3D
    
    For j = 1 To 21
        Rotate ori(Fra(j)).Z1, ori(Fra(j)).X1, 3.14159265258979 * 22.5 * ii / 180
        Rotate ori(Fra(j)).Z2, ori(Fra(j)).X2, 3.14159265258979 * 22.5 * ii / 180
        Rotate ori(Fra(j)).Z3, ori(Fra(j)).X3, 3.14159265258979 * 22.5 * ii / 180
        Rotate ori(Fra(j)).Z4, ori(Fra(j)).X4, 3.14159265258979 * 22.5 * ii / 180
    Next
    
    For i = 1 To 54
        Rotate ori(i).Z1, ori(i).X1, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z2, ori(i).X2, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z3, ori(i).X3, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z4, ori(i).X4, 3.14159265258979 * -Rott1.Value / 180
    Next
    
    For i = 1 To 54
        Rotate ori(i).Y1, ori(i).Z1, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y2, ori(i).Z2, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y3, ori(i).Z3, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y4, ori(i).Z4, 3.14159265258979 * Rott2.Value / 180
    Next
    
    DoEvents
    Picture1.Cls
    Picture1.Scale (-700, -700)-(300, 300)
    Call DrawCube
Next

Call RotateFace(CubeA, "d")
Call RotateFace(CubeA, "d")
Call RotateFace(CubeA, "d")
Call ViewCube(CubeA)

End Sub

Public Sub RotateLeftCWise()

Dim Fra(21) As Integer

Fra(1) = 28
Fra(2) = 29
Fra(3) = 30
Fra(4) = 31
Fra(5) = 32
Fra(6) = 33
Fra(7) = 34
Fra(8) = 35
Fra(9) = 36
Fra(10) = 1
Fra(11) = 4
Fra(12) = 7
Fra(13) = 37
Fra(14) = 40
Fra(15) = 43
Fra(16) = 46
Fra(17) = 49
Fra(18) = 52
Fra(19) = 21
Fra(20) = 24
Fra(21) = 27

For ii = 1 To 4
    
    Call Init3D
    
    For j = 1 To 21
        Rotate ori(Fra(j)).Y1, ori(Fra(j)).Z1, 3.14159265258979 * 22.5 * ii / 180
        Rotate ori(Fra(j)).Y2, ori(Fra(j)).Z2, 3.14159265258979 * 22.5 * ii / 180
        Rotate ori(Fra(j)).Y3, ori(Fra(j)).Z3, 3.14159265258979 * 22.5 * ii / 180
        Rotate ori(Fra(j)).Y4, ori(Fra(j)).Z4, 3.14159265258979 * 22.5 * ii / 180
    Next
    For i = 1 To 54
        Rotate ori(i).Z1, ori(i).X1, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z2, ori(i).X2, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z3, ori(i).X3, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z4, ori(i).X4, 3.14159265258979 * -Rott1.Value / 180
    Next
    For i = 1 To 54
        Rotate ori(i).Y1, ori(i).Z1, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y2, ori(i).Z2, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y3, ori(i).Z3, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y4, ori(i).Z4, 3.14159265258979 * Rott2.Value / 180
    Next
    
    DoEvents
    Picture1.Cls
    Picture1.Scale (-700, -700)-(300, 300)
    Call DrawCube

Next

Call RotateFace(CubeA, "l")
Call ViewCube(CubeA)
End Sub

Public Sub RotateLeftCCWise()

Dim Fra(21) As Integer

Fra(1) = 28
Fra(2) = 29
Fra(3) = 30
Fra(4) = 31
Fra(5) = 32
Fra(6) = 33
Fra(7) = 34
Fra(8) = 35
Fra(9) = 36
Fra(10) = 1
Fra(11) = 4
Fra(12) = 7
Fra(13) = 37
Fra(14) = 40
Fra(15) = 43
Fra(16) = 46
Fra(17) = 49
Fra(18) = 52
Fra(19) = 21
Fra(20) = 24
Fra(21) = 27

For ii = 1 To 4
    
    Call Init3D
    
    For j = 1 To 21
        Rotate ori(Fra(j)).Y1, ori(Fra(j)).Z1, 3.14159265258979 * -22.5 * ii / 180
        Rotate ori(Fra(j)).Y2, ori(Fra(j)).Z2, 3.14159265258979 * -22.5 * ii / 180
        Rotate ori(Fra(j)).Y3, ori(Fra(j)).Z3, 3.14159265258979 * -22.5 * ii / 180
        Rotate ori(Fra(j)).Y4, ori(Fra(j)).Z4, 3.14159265258979 * -22.5 * ii / 180
    Next
    For i = 1 To 54
        Rotate ori(i).Z1, ori(i).X1, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z2, ori(i).X2, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z3, ori(i).X3, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z4, ori(i).X4, 3.14159265258979 * -Rott1.Value / 180
    Next
    For i = 1 To 54
        Rotate ori(i).Y1, ori(i).Z1, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y2, ori(i).Z2, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y3, ori(i).Z3, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y4, ori(i).Z4, 3.14159265258979 * Rott2.Value / 180
    Next
    
    DoEvents
    Picture1.Cls
    Picture1.Scale (-700, -700)-(300, 300)
    Call DrawCube

Next

Call RotateFace(CubeA, "l")
Call RotateFace(CubeA, "l")
Call RotateFace(CubeA, "l")
Call ViewCube(CubeA)
End Sub

Public Sub RotateRightCWise()

Dim Fra(21) As Integer

Fra(1) = 10
Fra(2) = 11
Fra(3) = 12
Fra(4) = 13
Fra(5) = 14
Fra(6) = 15
Fra(7) = 16
Fra(8) = 17
Fra(9) = 18
Fra(10) = 3
Fra(11) = 6
Fra(12) = 9
Fra(13) = 19
Fra(14) = 22
Fra(15) = 25
Fra(16) = 39
Fra(17) = 42
Fra(18) = 45
Fra(19) = 48
Fra(20) = 51
Fra(21) = 54

For ii = 1 To 4
    
    Call Init3D
    
    For j = 1 To 21
        Rotate ori(Fra(j)).Y1, ori(Fra(j)).Z1, 3.14159265258979 * -22.5 * ii / 180
        Rotate ori(Fra(j)).Y2, ori(Fra(j)).Z2, 3.14159265258979 * -22.5 * ii / 180
        Rotate ori(Fra(j)).Y3, ori(Fra(j)).Z3, 3.14159265258979 * -22.5 * ii / 180
        Rotate ori(Fra(j)).Y4, ori(Fra(j)).Z4, 3.14159265258979 * -22.5 * ii / 180
    Next
    For i = 1 To 54
        Rotate ori(i).Z1, ori(i).X1, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z2, ori(i).X2, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z3, ori(i).X3, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z4, ori(i).X4, 3.14159265258979 * -Rott1.Value / 180
    Next
    For i = 1 To 54
        Rotate ori(i).Y1, ori(i).Z1, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y2, ori(i).Z2, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y3, ori(i).Z3, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y4, ori(i).Z4, 3.14159265258979 * Rott2.Value / 180
    Next
    
    DoEvents
    Picture1.Cls
    Picture1.Scale (-700, -700)-(300, 300)
    Call DrawCube

Next

Call RotateFace(CubeA, "r")
Call ViewCube(CubeA)

End Sub

Public Sub RotateRightCCWise()

Dim Fra(21) As Integer

Fra(1) = 10
Fra(2) = 11
Fra(3) = 12
Fra(4) = 13
Fra(5) = 14
Fra(6) = 15
Fra(7) = 16
Fra(8) = 17
Fra(9) = 18
Fra(10) = 3
Fra(11) = 6
Fra(12) = 9
Fra(13) = 19
Fra(14) = 22
Fra(15) = 25
Fra(16) = 39
Fra(17) = 42
Fra(18) = 45
Fra(19) = 48
Fra(20) = 51
Fra(21) = 54

For ii = 1 To 4

    Call Init3D
    
    For j = 1 To 21
        Rotate ori(Fra(j)).Y1, ori(Fra(j)).Z1, 3.14159265258979 * 22.5 * ii / 180
        Rotate ori(Fra(j)).Y2, ori(Fra(j)).Z2, 3.14159265258979 * 22.5 * ii / 180
        Rotate ori(Fra(j)).Y3, ori(Fra(j)).Z3, 3.14159265258979 * 22.5 * ii / 180
        Rotate ori(Fra(j)).Y4, ori(Fra(j)).Z4, 3.14159265258979 * 22.5 * ii / 180
    Next
    
    For i = 1 To 54
        Rotate ori(i).Z1, ori(i).X1, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z2, ori(i).X2, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z3, ori(i).X3, 3.14159265258979 * -Rott1.Value / 180
        Rotate ori(i).Z4, ori(i).X4, 3.14159265258979 * -Rott1.Value / 180
    Next
    
    For i = 1 To 54
        Rotate ori(i).Y1, ori(i).Z1, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y2, ori(i).Z2, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y3, ori(i).Z3, 3.14159265258979 * Rott2.Value / 180
        Rotate ori(i).Y4, ori(i).Z4, 3.14159265258979 * Rott2.Value / 180
    Next
    
    DoEvents
    
    Picture1.Cls
    Picture1.Scale (-700, -700)-(300, 300)
    Call DrawCube
    
Next

Call RotateFace(CubeA, "r")
Call RotateFace(CubeA, "r")
Call RotateFace(CubeA, "r")
Call ViewCube(CubeA)

End Sub

Public Sub ScrambleCube()

For i = 1 To Val(Text1.Text)
    
    Randomize
    n = Round(Rnd * 2342123) Mod 12
    
    Select Case (n + 1)
        
        Case 1
            Call RotateFace(CubeA, "f")
        
        Case 2
            Call RotateFace(CubeA, "f")
            Call RotateFace(CubeA, "f")
            Call RotateFace(CubeA, "f")
        
        Case 3
            Call RotateFace(CubeA, "r")
        
        Case 4
            Call RotateFace(CubeA, "r")
            Call RotateFace(CubeA, "r")
            Call RotateFace(CubeA, "r")
    
        Case 5
            Call RotateFace(CubeA, "l")
        
        Case 6
            Call RotateFace(CubeA, "l")
            Call RotateFace(CubeA, "l")
            Call RotateFace(CubeA, "l")

        Case 7
            Call RotateFace(CubeA, "b")
    
        Case 8
            Call RotateFace(CubeA, "b")
            Call RotateFace(CubeA, "b")
            Call RotateFace(CubeA, "b")

        Case 9
            Call RotateFace(CubeA, "t")
        
        Case 10
            Call RotateFace(CubeA, "t")
            Call RotateFace(CubeA, "t")
            Call RotateFace(CubeA, "t")

        Case 11
            Call RotateFace(CubeA, "d")
        
        Case 12
            Call RotateFace(CubeA, "d")
            Call RotateFace(CubeA, "d")
            Call RotateFace(CubeA, "d")

    End Select

Next

DoEvents
Call ViewCube(CubeA)

End Sub

Public Sub ResetCube()

CubeA = "RRRRRRRRRYYYYYYYYYPPPPPPPPPWWWWWWWWWBBBBBBBBBGGGGGGGGG"

Rott1.Value = -30
Rott2.Value = 30

Call Rott1_Change

Call ViewCube(CubeA)

End Sub

Private Sub Rotate(X, Y, q)

xd = X * Cos(q) - Y * Sin(q)
yd = X * Sin(q) + Y * Cos(q)

X = xd
Y = yd

End Sub

Private Sub DrawPolygon(Fram As Frames, Colour As Long)
On Error Resume Next

X1 = Fram.X1
Y1 = Fram.Y1
Z1 = Fram.Z1
X2 = Fram.X2
Y2 = Fram.Y2
Z2 = Fram.Z2
X3 = Fram.X3
Y3 = Fram.Y3
Z3 = Fram.Z3
X4 = Fram.X4
Y4 = Fram.Y4
Z4 = Fram.Z4

X1 = X1 * (1000 - Z1) / 1000
Y1 = Y1 * (1000 - Z1) / 1000
X2 = X2 * (1000 - Z2) / 1000
Y2 = Y2 * (1000 - Z2) / 1000
X3 = X3 * (1000 - Z3) / 1000
Y3 = Y3 * (1000 - Z3) / 1000
X4 = X4 * (1000 - Z4) / 1000
Y4 = Y4 * (1000 - Z4) / 1000

For xx = X1 To X2 Step Sgn(X2 - X1) * Abs(X2 - X1) / 30 + 0.0000001
    
    Picture1.Line (xx, Y1 + (Y2 - Y1) * (xx - X1) / _
                    (X2 - X1))-(X4 + (X3 - X4) * (xx - X1) / _
                    (X2 - X1), Y4 + (Y3 - Y4) * (xx - X1) / _
                    (X2 - X1)), Colour
Next

Picture1.Line (X1, Y1)-(X2, Y2), 0
Picture1.Line (X2, Y2)-(X3, Y3), 0
Picture1.Line (X3, Y3)-(X4, Y4), 0
Picture1.Line (X4, Y4)-(X1, Y1), 0

End Sub

Private Sub DrawPolygon1(Fram As Frames, Colour As Long)
On Error Resume Next

X1 = Fram.X1
Y1 = Fram.Y1
Z1 = -Fram.Z1 + 300

X2 = Fram.X2
Y2 = Fram.Y2
Z2 = -Fram.Z2 + 300

X3 = Fram.X3
Y3 = Fram.Y3
Z3 = -Fram.Z3 + 300

X4 = Fram.X4
Y4 = Fram.Y4
Z4 = -Fram.Z4 + 300

X1 = X1 * (1000 - Z1) / 1000
Y1 = Y1 * (1000 - Z1) / 1000
X2 = X2 * (1000 - Z2) / 1000
Y2 = Y2 * (1000 - Z2) / 1000

X3 = X3 * (1000 - Z3) / 1000
Y3 = Y3 * (1000 - Z3) / 1000
X4 = X4 * (1000 - Z4) / 1000
Y4 = Y4 * (1000 - Z4) / 1000

For xx = X1 To X2 Step Sgn(X2 - X1) * Abs(X2 - X1) / 30 + 0.0000001
    Picture1.Line (xx, Y1 + (Y2 - Y1) * (xx - X1) / _
                    (X2 - X1))-(X4 + (X3 - X4) * (xx - X1) / _
                    (X2 - X1), Y4 + (Y3 - Y4) * (xx - X1) / _
                    (X2 - X1)), Colour
Next

Picture1.Line (X1, Y1)-(X2, Y2), 0
Picture1.Line (X2, Y2)-(X3, Y3), 0
Picture1.Line (X3, Y3)-(X4, Y4), 0
Picture1.Line (X4, Y4)-(X1, Y1), 0

End Sub

Private Sub RotateFace(Cube As String, Face As String)

Select Case Face
    
    Case "f" ' Rotate front face Clock-wise
        
        temp$ = Mid(Cube, 2, 1)
        
        Mid(Cube, 2, 1) = Mid(Cube, 4, 1)
        Mid(Cube, 4, 1) = Mid(Cube, 8, 1)
        Mid(Cube, 8, 1) = Mid(Cube, 6, 1)
        Mid(Cube, 6, 1) = temp$
        
        temp$ = Mid(Cube, 1, 1)

        Mid(Cube, 1, 1) = Mid(Cube, 7, 1)
        Mid(Cube, 7, 1) = Mid(Cube, 9, 1)
        Mid(Cube, 9, 1) = Mid(Cube, 3, 1)
        Mid(Cube, 3, 1) = temp$

        temp$ = Mid(Cube, 43, 1)

        Mid(Cube, 43, 1) = Mid(Cube, 36, 1)
        Mid(Cube, 36, 1) = Mid(Cube, 54, 1)
        Mid(Cube, 54, 1) = Mid(Cube, 10, 1)
        Mid(Cube, 10, 1) = temp$

        temp$ = Mid(Cube, 44, 1)

        Mid(Cube, 44, 1) = Mid(Cube, 33, 1)
        Mid(Cube, 33, 1) = Mid(Cube, 53, 1)
        Mid(Cube, 53, 1) = Mid(Cube, 13, 1)
        Mid(Cube, 13, 1) = temp$

        temp$ = Mid(Cube, 45, 1)

        Mid(Cube, 45, 1) = Mid(Cube, 30, 1)
        Mid(Cube, 30, 1) = Mid(Cube, 52, 1)
        Mid(Cube, 52, 1) = Mid(Cube, 16, 1)
        Mid(Cube, 16, 1) = temp$

    Case "r" ' Rotate front face Clock-wise
        
        temp$ = Mid(Cube, 11, 1)

        Mid(Cube, 11, 1) = Mid(Cube, 13, 1)
        Mid(Cube, 13, 1) = Mid(Cube, 17, 1)
        Mid(Cube, 17, 1) = Mid(Cube, 15, 1)
        Mid(Cube, 15, 1) = temp$
        
        temp$ = Mid(Cube, 10, 1)
    
        Mid(Cube, 10, 1) = Mid(Cube, 16, 1)
        Mid(Cube, 16, 1) = Mid(Cube, 18, 1)
        Mid(Cube, 18, 1) = Mid(Cube, 12, 1)
        Mid(Cube, 12, 1) = temp$

        temp$ = Mid(Cube, 45, 1)

        Mid(Cube, 45, 1) = Mid(Cube, 9, 1)
        Mid(Cube, 9, 1) = Mid(Cube, 48, 1)
        Mid(Cube, 48, 1) = Mid(Cube, 19, 1)
        Mid(Cube, 19, 1) = temp$

        temp$ = Mid(Cube, 42, 1)

        Mid(Cube, 42, 1) = Mid(Cube, 6, 1)
        Mid(Cube, 6, 1) = Mid(Cube, 51, 1)
        Mid(Cube, 51, 1) = Mid(Cube, 22, 1)
        Mid(Cube, 22, 1) = temp$

        temp$ = Mid(Cube, 39, 1)
        
        Mid(Cube, 39, 1) = Mid(Cube, 3, 1)
        Mid(Cube, 3, 1) = Mid(Cube, 54, 1)
        Mid(Cube, 54, 1) = Mid(Cube, 25, 1)
        Mid(Cube, 25, 1) = temp$

    Case "l" ' Rotate front face Clock-wise

        temp$ = Mid(Cube, 30, 1)
        
        Mid(Cube, 30, 1) = Mid(Cube, 28, 1)
        Mid(Cube, 28, 1) = Mid(Cube, 34, 1)
        Mid(Cube, 34, 1) = Mid(Cube, 36, 1)
        Mid(Cube, 36, 1) = temp$
        
        temp$ = Mid(Cube, 29, 1)

        Mid(Cube, 29, 1) = Mid(Cube, 31, 1)
        Mid(Cube, 31, 1) = Mid(Cube, 35, 1)
        Mid(Cube, 35, 1) = Mid(Cube, 33, 1)
        Mid(Cube, 33, 1) = temp$

        temp$ = Mid(Cube, 1, 1)

        Mid(Cube, 1, 1) = Mid(Cube, 37, 1)
        Mid(Cube, 37, 1) = Mid(Cube, 27, 1)
        Mid(Cube, 27, 1) = Mid(Cube, 52, 1)
        Mid(Cube, 52, 1) = temp$

        temp$ = Mid(Cube, 4, 1)

        Mid(Cube, 4, 1) = Mid(Cube, 40, 1)
        Mid(Cube, 40, 1) = Mid(Cube, 24, 1)
        Mid(Cube, 24, 1) = Mid(Cube, 49, 1)
        Mid(Cube, 49, 1) = temp$

        temp$ = Mid(Cube, 7, 1)

        Mid(Cube, 7, 1) = Mid(Cube, 43, 1)
        Mid(Cube, 43, 1) = Mid(Cube, 21, 1)
        Mid(Cube, 21, 1) = Mid(Cube, 46, 1)
        Mid(Cube, 46, 1) = temp$

    Case "t" ' Rotate front face Clock-wise

        temp$ = Mid(Cube, 37, 1)
        
        Mid(Cube, 37, 1) = Mid(Cube, 43, 1)
        Mid(Cube, 43, 1) = Mid(Cube, 45, 1)
        Mid(Cube, 45, 1) = Mid(Cube, 39, 1)
        Mid(Cube, 39, 1) = temp$

        temp$ = Mid(Cube, 40, 1)

        Mid(Cube, 40, 1) = Mid(Cube, 44, 1)
        Mid(Cube, 44, 1) = Mid(Cube, 42, 1)
        Mid(Cube, 42, 1) = Mid(Cube, 38, 1)
        Mid(Cube, 38, 1) = temp$

        temp$ = Mid(Cube, 1, 1)

        Mid(Cube, 1, 1) = Mid(Cube, 10, 1)
        Mid(Cube, 10, 1) = Mid(Cube, 19, 1)
        Mid(Cube, 19, 1) = Mid(Cube, 28, 1)
        Mid(Cube, 28, 1) = temp$

        temp$ = Mid(Cube, 29, 1)

        Mid(Cube, 29, 1) = Mid(Cube, 2, 1)
        Mid(Cube, 2, 1) = Mid(Cube, 11, 1)
        Mid(Cube, 11, 1) = Mid(Cube, 20, 1)
        Mid(Cube, 20, 1) = temp$

        temp$ = Mid(Cube, 3, 1)
    
        Mid(Cube, 3, 1) = Mid(Cube, 12, 1)
        Mid(Cube, 12, 1) = Mid(Cube, 21, 1)
        Mid(Cube, 21, 1) = Mid(Cube, 30, 1)
        Mid(Cube, 30, 1) = temp$

    Case "b" ' Rotate front face Clock-wise

        temp$ = Mid(Cube, 21, 1)
        
        Mid(Cube, 21, 1) = Mid(Cube, 19, 1)
        Mid(Cube, 19, 1) = Mid(Cube, 25, 1)
        Mid(Cube, 25, 1) = Mid(Cube, 27, 1)
        Mid(Cube, 27, 1) = temp$

        temp$ = Mid(Cube, 20, 1)

        Mid(Cube, 20, 1) = Mid(Cube, 22, 1)
        Mid(Cube, 22, 1) = Mid(Cube, 26, 1)
        Mid(Cube, 26, 1) = Mid(Cube, 24, 1)
        Mid(Cube, 24, 1) = temp$

        temp$ = Mid(Cube, 37, 1)

        Mid(Cube, 37, 1) = Mid(Cube, 12, 1)
        Mid(Cube, 12, 1) = Mid(Cube, 48, 1)
        Mid(Cube, 48, 1) = Mid(Cube, 34, 1)
        Mid(Cube, 34, 1) = temp$

        temp$ = Mid(Cube, 38, 1)

        Mid(Cube, 38, 1) = Mid(Cube, 15, 1)
        Mid(Cube, 15, 1) = Mid(Cube, 47, 1)
        Mid(Cube, 47, 1) = Mid(Cube, 31, 1)
        Mid(Cube, 31, 1) = temp$

        temp$ = Mid(Cube, 39, 1)

        Mid(Cube, 39, 1) = Mid(Cube, 18, 1)
        Mid(Cube, 18, 1) = Mid(Cube, 46, 1)
        Mid(Cube, 46, 1) = Mid(Cube, 28, 1)
        Mid(Cube, 28, 1) = temp$

    Case "d" ' Rotate front face Clock-wise
    
        temp$ = Mid(Cube, 46, 1)
        
        Mid(Cube, 46, 1) = Mid(Cube, 48, 1)
        Mid(Cube, 48, 1) = Mid(Cube, 54, 1)
        Mid(Cube, 54, 1) = Mid(Cube, 52, 1)
        Mid(Cube, 52, 1) = temp$

        temp$ = Mid(Cube, 47, 1)

        Mid(Cube, 47, 1) = Mid(Cube, 51, 1)
        Mid(Cube, 51, 1) = Mid(Cube, 53, 1)
        Mid(Cube, 53, 1) = Mid(Cube, 49, 1)
        Mid(Cube, 49, 1) = temp$

        temp$ = Mid(Cube, 16, 1)

        Mid(Cube, 16, 1) = Mid(Cube, 7, 1)
        Mid(Cube, 7, 1) = Mid(Cube, 34, 1)
        Mid(Cube, 34, 1) = Mid(Cube, 25, 1)
        Mid(Cube, 25, 1) = temp$

        temp$ = Mid(Cube, 8, 1)
        
        Mid(Cube, 8, 1) = Mid(Cube, 35, 1)
        Mid(Cube, 35, 1) = Mid(Cube, 26, 1)
        Mid(Cube, 26, 1) = Mid(Cube, 17, 1)
        Mid(Cube, 17, 1) = temp$

        temp$ = Mid(Cube, 9, 1)

        Mid(Cube, 9, 1) = Mid(Cube, 36, 1)
        Mid(Cube, 36, 1) = Mid(Cube, 27, 1)
        Mid(Cube, 27, 1) = Mid(Cube, 18, 1)
        Mid(Cube, 18, 1) = temp$

End Select

End Sub

Private Sub lblRotate_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Timer3.Enabled = True
RotateVal = Index

End Sub

Private Sub lblRotate_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Timer3.Enabled = False

End Sub

Private Sub Timer3_Timer()

Select Case RotateVal

    'up
    Case 0
                
        RotateSgn = "-"
        If Rott2.Value = -150 Then Exit Sub
        Rott2.Value = Rott2.Value - 30
    
    'down
    Case 1
        
        RotateSgn = "+"
        If Rott2.Value = 150 Then Exit Sub
        Rott2.Value = Rott2.Value + 30
    
    'right
    Case 2
        
        RotateSgn = "+"
        If Rott1.Value = 180 Then Exit Sub
        Rott1.Value = Rott1.Value + 30
    
    'left
    Case 3
        
        RotateSgn = "-"
        If Rott1.Value = -180 Then Exit Sub
        Rott1.Value = Rott1.Value - 30

End Select

End Sub

Private Sub UserControl_Initialize()

CubeA = "RRRRRRRRRYYYYYYYYYPPPPPPPPPWWWWWWWWWBBBBBBBBBGGGGGGGGG"

Picture1.Scale (-700, -700)-(300, 300)

Call GetCoordinates
Call Rott1_Change
Call Rott2_Change

End Sub

Public Function GetXCoor() As Integer

GetXCoor = Rott1.Value

End Function

Public Function GetYCoor() As Integer

GetYCoor = Rott2.Value

End Function

Public Function GetCube() As String

GetCube = CubeA

End Function

Public Sub GetRotation(Coor As String, Pos As String, TypeOfRotate As String)
Dim rubik

rubik = ""
rubik = Coor & "," & Pos & "*"

For i = 0 To List1.ListCount - 1
    If List1.List(i) Like rubik Then
        GetMove List1.List(i)
        Rotation Movement, TypeOfRotate
        Exit For
    End If
Next

End Sub

Sub GetCoordinates()

List1.AddItem "-30/150,L-L"
List1.AddItem "-30/150,R-R"
List1.AddItem "-30/150,T-BO"
List1.AddItem "-30/150,BO-T"
List1.AddItem "-30/150,F-BA"
List1.AddItem "-30/150,BA-F"

List1.AddItem "-30/120,L-L"
List1.AddItem "-30/120,R-R"
List1.AddItem "-30/120,T-BA"
List1.AddItem "-30/120,BO-F"
List1.AddItem "-30/120,F-T"
List1.AddItem "-30/120,BA-BO"

List1.AddItem "-30/60,L-L"
List1.AddItem "-30/60,R-R"
List1.AddItem "-30/60,T-BA"
List1.AddItem "-30/60,BO-F"
List1.AddItem "-30/60,F-T"
List1.AddItem "-30/60,BA-BO"

List1.AddItem "-30/30,L-L"
List1.AddItem "-30/30,R-R"
List1.AddItem "-30/30,T-T"
List1.AddItem "-30/30,BO-BO"
List1.AddItem "-30/30,F-F"
List1.AddItem "-30/30,BA-BA"

List1.AddItem "-30/-30,L-L"
List1.AddItem "-30/-30,R-R"
List1.AddItem "-30/-30,T-T"
List1.AddItem "-30/-30,BO-BO"
List1.AddItem "-30/-30,F-F"
List1.AddItem "-30/-30,BA-BA"

List1.AddItem "-30/-60,L-L"
List1.AddItem "-30/-60,R-R"
List1.AddItem "-30/-60,T-F"
List1.AddItem "-30/-60,BO-BA"
List1.AddItem "-30/-60,F-BO"
List1.AddItem "-30/-60,BA-T"

List1.AddItem "-30/-120,L-L"
List1.AddItem "-30/-120,R-R"
List1.AddItem "-30/-120,T-F"
List1.AddItem "-30/-120,BO-BA"
List1.AddItem "-30/-120,F-BO"
List1.AddItem "-30/-120,BA-T"

List1.AddItem "-30/-150,L-L"
List1.AddItem "-30/-150,R-R"
List1.AddItem "-30/-150,T-BO"
List1.AddItem "-30/-150,BO-T"
List1.AddItem "-30/-150,F-BA"
List1.AddItem "-30/-150,BA-F"

List1.AddItem "0/150,L-L"
List1.AddItem "0/150,R-R"
List1.AddItem "0/150,T-BO"
List1.AddItem "0/150,BO-T"
List1.AddItem "0/150,F-BA"
List1.AddItem "0/150,BA-F"

List1.AddItem "0/120,L-L"
List1.AddItem "0/120,R-R"
List1.AddItem "0/120,T-BA"
List1.AddItem "0/120,BO-F"
List1.AddItem "0/120,F-T"
List1.AddItem "0/120,BA-BO"

List1.AddItem "0/60,L-L"
List1.AddItem "0/60,R-R"
List1.AddItem "0/60,T-BA"
List1.AddItem "0/60,BO-F"
List1.AddItem "0/60,F-T"
List1.AddItem "0/60,BA-BO"

List1.AddItem "0/30,L-L"
List1.AddItem "0/30,R-R"
List1.AddItem "0/30,T-T"
List1.AddItem "0/30,BO-BO"
List1.AddItem "0/30,F-F"
List1.AddItem "0/30,BA-BA"

List1.AddItem "0/-30,L-L"
List1.AddItem "0/-30,R-R"
List1.AddItem "0/-30,T-T"
List1.AddItem "0/-30,BO-BO"
List1.AddItem "0/-30,F-F"
List1.AddItem "0/-30,BA-BA"

List1.AddItem "0/-60,L-L"
List1.AddItem "0/-60,R-R"
List1.AddItem "0/-60,T-F"
List1.AddItem "0/-60,BO-BA"
List1.AddItem "0/-60,F-BO"
List1.AddItem "0/-60,BA-T"

List1.AddItem "0/-120,L-L"
List1.AddItem "0/-120,R-R"
List1.AddItem "0/-120,T-F"
List1.AddItem "0/-120,BO-BA"
List1.AddItem "0/-120,F-BO"
List1.AddItem "0/-120,BA-T"

List1.AddItem "0/-150,L-L"
List1.AddItem "0/-150,R-R"
List1.AddItem "0/-150,T-BO"
List1.AddItem "0/-150,BO-T"
List1.AddItem "0/-150,F-BA"
List1.AddItem "0/-150,BA-F"

List1.AddItem "30/150,L-L"
List1.AddItem "30/150,R-R"
List1.AddItem "30/150,T-BO"
List1.AddItem "30/150,BO-T"
List1.AddItem "30/150,F-BA"
List1.AddItem "30/150,BA-F"

List1.AddItem "30/120,L-L"
List1.AddItem "30/120,R-R"
List1.AddItem "30/120,T-BA"
List1.AddItem "30/120,BO-F"
List1.AddItem "30/120,F-T"
List1.AddItem "30/120,BA-BO"

List1.AddItem "30/60,L-L"
List1.AddItem "30/60,R-R"
List1.AddItem "30/60,T-BA"
List1.AddItem "30/60,BO-F"
List1.AddItem "30/60,F-T"
List1.AddItem "30/60,BA-BO"

List1.AddItem "30/30,L-L"
List1.AddItem "30/30,R-R"
List1.AddItem "30/30,T-T"
List1.AddItem "30/30,BO-BO"
List1.AddItem "30/30,F-F"
List1.AddItem "30/30,BA-BA"

List1.AddItem "30/-30,L-L"
List1.AddItem "30/-30,R-R"
List1.AddItem "30/-30,T-T"
List1.AddItem "30/-30,BO-BO"
List1.AddItem "30/-30,F-F"
List1.AddItem "30/-30,BA-BA"

List1.AddItem "30/-60,L-L"
List1.AddItem "30/-60,R-R"
List1.AddItem "30/-60,T-F"
List1.AddItem "30/-60,BO-BA"
List1.AddItem "30/-60,F-BO"
List1.AddItem "30/-60,BA-T"

List1.AddItem "30/-120,L-L"
List1.AddItem "30/-120,R-R"
List1.AddItem "30/-120,T-F"
List1.AddItem "30/-120,BO-BA"
List1.AddItem "30/-120,F-BO"
List1.AddItem "30/-120,BA-T"

List1.AddItem "30/-150,L-L"
List1.AddItem "30/-150,R-R"
List1.AddItem "30/-150,T-BO"
List1.AddItem "30/-150,BO-T"
List1.AddItem "30/-150,F-BA"
List1.AddItem "30/-150,BA-F"

List1.AddItem "60/150,L-BA"
List1.AddItem "60/150,R-F"
List1.AddItem "60/150,T-BO"
List1.AddItem "60/150,BO-T"
List1.AddItem "60/150,F-R"
List1.AddItem "60/150,BA-L"

List1.AddItem "60/120,L-BA"
List1.AddItem "60/120,R-F"
List1.AddItem "60/120,T-R"
List1.AddItem "60/120,BO-L"
List1.AddItem "60/120,F-T"
List1.AddItem "60/120,BA-BO"

List1.AddItem "60/60,L-BA"
List1.AddItem "60/60,R-F"
List1.AddItem "60/60,T-R"
List1.AddItem "60/60,BO-L"
List1.AddItem "60/60,F-T"
List1.AddItem "60/60,BA-BO"

List1.AddItem "60/30,L-BA"
List1.AddItem "60/30,R-F"
List1.AddItem "60/30,T-T"
List1.AddItem "60/30,BO-BO"
List1.AddItem "60/30,F-L"
List1.AddItem "60/30,BA-R"

List1.AddItem "60/-30,L-BA"
List1.AddItem "60/-30,R-F"
List1.AddItem "60/-30,T-T"
List1.AddItem "60/-30,BO-BO"
List1.AddItem "60/-30,F-L"
List1.AddItem "60/-30,BA-R"

List1.AddItem "60/-60,L-BA"
List1.AddItem "60/-60,R-F"
List1.AddItem "60/-60,T-L"
List1.AddItem "60/-60,BO-R"
List1.AddItem "60/-60,F-BO"
List1.AddItem "60/-60,BA-T"

List1.AddItem "60/-120,L-BA"
List1.AddItem "60/-120,R-F"
List1.AddItem "60/-120,T-L"
List1.AddItem "60/-120,BO-R"
List1.AddItem "60/-120,F-BO"
List1.AddItem "60/-120,BA-T"

List1.AddItem "60/-150,L-BA"
List1.AddItem "60/-150,R-F"
List1.AddItem "60/-150,T-BO"
List1.AddItem "60/-150,BO-T"
List1.AddItem "60/-150,F-R"
List1.AddItem "60/-150,BA-L"

List1.AddItem "120/150,L-F"
List1.AddItem "120/150,R-BA"
List1.AddItem "120/150,T-BO"
List1.AddItem "120/150,BO-T"
List1.AddItem "120/150,F-R"
List1.AddItem "120/150,BA-L"

List1.AddItem "120/120,L-BA"
List1.AddItem "120/120,R-F"
List1.AddItem "120/120,T-R"
List1.AddItem "120/120,BO-L"
List1.AddItem "120/120,F-T"
List1.AddItem "120/120,BA-BO"

List1.AddItem "120/60,L-BA"
List1.AddItem "120/60,R-F"
List1.AddItem "120/60,T-R"
List1.AddItem "120/60,BO-L"
List1.AddItem "120/60,F-T"
List1.AddItem "120/60,BA-BO"

List1.AddItem "120/30,L-BA"
List1.AddItem "120/30,R-F"
List1.AddItem "120/30,T-T"
List1.AddItem "120/30,BO-BO"
List1.AddItem "120/30,F-L"
List1.AddItem "120/30,BA-R"

List1.AddItem "120/-30,L-BA"
List1.AddItem "120/-30,R-F"
List1.AddItem "120/-30,T-T"
List1.AddItem "120/-30,BO-BO"
List1.AddItem "120/-30,F-L"
List1.AddItem "120/-30,BA-R"

List1.AddItem "120/-60,L-BA"
List1.AddItem "120/-60,R-F"
List1.AddItem "120/-60,T-L"
List1.AddItem "120/-60,BO-R"
List1.AddItem "120/-60,F-BO"
List1.AddItem "120/-60,BA-T"

List1.AddItem "120/-120,L-BA"
List1.AddItem "120/-120,R-F"
List1.AddItem "120/-120,T-L"
List1.AddItem "120/-120,BO-R"
List1.AddItem "120/-120,F-BO"
List1.AddItem "120/-120,BA-T"

List1.AddItem "120/-150,L-BA"
List1.AddItem "120/-150,R-F"
List1.AddItem "120/-150,T-BO"
List1.AddItem "120/-150,BO-T"
List1.AddItem "120/-150,F-R"
List1.AddItem "120/-150,BA-L"

List1.AddItem "150/150,L-R"
List1.AddItem "150/150,R-L"
List1.AddItem "150/150,T-BO"
List1.AddItem "150/150,BO-T"
List1.AddItem "150/150,F-F"
List1.AddItem "150/150,BA-BA"

List1.AddItem "150/120,L-R"
List1.AddItem "150/120,R-L"
List1.AddItem "150/120,T-F"
List1.AddItem "150/120,BO-BA"
List1.AddItem "150/120,F-T"
List1.AddItem "150/120,BA-BO"

List1.AddItem "150/60,L-R"
List1.AddItem "150/60,R-L"
List1.AddItem "150/60,T-F"
List1.AddItem "150/60,BO-BA"
List1.AddItem "150/60,F-T"
List1.AddItem "150/60,BA-BO"

List1.AddItem "150/30,L-R"
List1.AddItem "150/30,R-L"
List1.AddItem "150/30,T-T"
List1.AddItem "150/30,BO-BO"
List1.AddItem "150/30,F-BA"
List1.AddItem "150/30,BA-F"

List1.AddItem "150/-30,L-R"
List1.AddItem "150/-30,R-L"
List1.AddItem "150/-30,T-T"
List1.AddItem "150/-30,BO-BO"
List1.AddItem "150/-30,F-BA"
List1.AddItem "150/-30,BA-F"

List1.AddItem "150/-60,L-R"
List1.AddItem "150/-60,R-L"
List1.AddItem "150/-60,T-BA"
List1.AddItem "150/-60,BO-F"
List1.AddItem "150/-60,F-BO"
List1.AddItem "150/-60,BA-T"

List1.AddItem "150/-120,L-R"
List1.AddItem "150/-120,R-L"
List1.AddItem "150/-120,T-BA"
List1.AddItem "150/-120,BO-F"
List1.AddItem "150/-120,F-BO"
List1.AddItem "150/-120,BA-T"

List1.AddItem "150/-150,L-R"
List1.AddItem "150/-150,R-L"
List1.AddItem "150/-150,T-BO"
List1.AddItem "150/-150,BO-T"
List1.AddItem "150/-150,F-F"
List1.AddItem "150/-150,BA-BA"

List1.AddItem "180/150,L-R"
List1.AddItem "180/150,R-L"
List1.AddItem "180/150,T-BO"
List1.AddItem "180/150,BO-T"
List1.AddItem "180/150,F-F"
List1.AddItem "180/150,BA-BA"

List1.AddItem "180/120,L-R"
List1.AddItem "180/120,R-L"
List1.AddItem "180/120,T-F"
List1.AddItem "180/120,BO-BA"
List1.AddItem "180/120,F-T"
List1.AddItem "180/120,BA-BO"

List1.AddItem "180/60,L-R"
List1.AddItem "180/60,R-L"
List1.AddItem "180/60,T-F"
List1.AddItem "180/60,BO-BA"
List1.AddItem "180/60,F-T"
List1.AddItem "180/60,BA-BO"

List1.AddItem "180/30,L-R"
List1.AddItem "180/30,R-L"
List1.AddItem "180/30,T-T"
List1.AddItem "180/30,BO-BO"
List1.AddItem "180/30,F-BA"
List1.AddItem "180/30,BA-F"

List1.AddItem "180/-30,L-R"
List1.AddItem "180/-30,R-L"
List1.AddItem "180/-30,T-T"
List1.AddItem "180/-30,BO-BO"
List1.AddItem "180/-30,F-BA"
List1.AddItem "180/-30,BA-F"

List1.AddItem "180/-60,L-R"
List1.AddItem "180/-60,R-L"
List1.AddItem "180/-60,T-BA"
List1.AddItem "180/-60,BO-F"
List1.AddItem "180/-60,F-BO"
List1.AddItem "180/-60,BA-T"

List1.AddItem "180/-120,L-R"
List1.AddItem "180/-120,R-L"
List1.AddItem "180/-120,T-BA"
List1.AddItem "180/-120,BO-F"
List1.AddItem "180/-120,F-BO"
List1.AddItem "180/-120,BA-T"

List1.AddItem "180/-150,L-R"
List1.AddItem "180/-150,R-L"
List1.AddItem "180/-150,T-BO"
List1.AddItem "180/-150,BO-T"
List1.AddItem "180/-150,F-F"
List1.AddItem "180/-150,BA-BA"

List1.AddItem "-60/150,L-F"
List1.AddItem "-60/150,R-BA"
List1.AddItem "-60/150,T-BO"
List1.AddItem "-60/150,BO-T"
List1.AddItem "-60/150,F-L"
List1.AddItem "-60/150,BA-R"

List1.AddItem "-60/120,L-F"
List1.AddItem "-60/120,R-BA"
List1.AddItem "-60/120,T-L"
List1.AddItem "-60/120,BO-R"
List1.AddItem "-60/120,F-T"
List1.AddItem "-60/120,BA-BO"

List1.AddItem "-60/60,L-F"
List1.AddItem "-60/60,R-BA"
List1.AddItem "-60/60,T-L"
List1.AddItem "-60/60,BO-R"
List1.AddItem "-60/60,F-T"
List1.AddItem "-60/60,BA-BO"

List1.AddItem "-60/30,L-F"
List1.AddItem "-60/30,R-BA"
List1.AddItem "-60/30,T-T"
List1.AddItem "-60/30,BO-BO"
List1.AddItem "-60/30,F-R"
List1.AddItem "-60/30,BA-L"

List1.AddItem "-60/-30,L-F"
List1.AddItem "-60/-30,R-BA"
List1.AddItem "-60/-30,T-T"
List1.AddItem "-60/-30,BO-BO"
List1.AddItem "-60/-30,F-R"
List1.AddItem "-60/-30,BA-L"

List1.AddItem "-60/-60,L-F"
List1.AddItem "-60/-60,R-BA"
List1.AddItem "-60/-60,T-R"
List1.AddItem "-60/-60,BO-L"
List1.AddItem "-60/-60,F-BO"
List1.AddItem "-60/-60,BA-T"

List1.AddItem "-60/-120,L-F"
List1.AddItem "-60/-120,R-BA"
List1.AddItem "-60/-120,T-R"
List1.AddItem "-60/-120,BO-L"
List1.AddItem "-60/-120,F-BO"
List1.AddItem "-60/-120,BA-T"

List1.AddItem "-60/-150,L-F"
List1.AddItem "-60/-150,R-BA"
List1.AddItem "-60/-150,T-BO"
List1.AddItem "-60/-150,BO-T"
List1.AddItem "-60/-150,F-L"
List1.AddItem "-60/-150,BA-R"

List1.AddItem "-120/150,L-F"
List1.AddItem "-120/150,R-BA"
List1.AddItem "-120/150,T-BO"
List1.AddItem "-120/150,BO-T"
List1.AddItem "-120/150,F-L"
List1.AddItem "-120/150,BA-R"

List1.AddItem "-120/120,L-F"
List1.AddItem "-120/120,R-BA"
List1.AddItem "-120/120,T-L"
List1.AddItem "-120/120,BO-R"
List1.AddItem "-120/120,F-T"
List1.AddItem "-120/120,BA-BO"

List1.AddItem "-120/60,L-F"
List1.AddItem "-120/60,R-BA"
List1.AddItem "-120/60,T-L"
List1.AddItem "-120/60,BO-R"
List1.AddItem "-120/60,F-T"
List1.AddItem "-120/60,BA-BO"

List1.AddItem "-120/30,L-F"
List1.AddItem "-120/30,R-BA"
List1.AddItem "-120/30,T-T"
List1.AddItem "-120/30,BO-BO"
List1.AddItem "-120/30,F-R"
List1.AddItem "-120/30,BA-L"

List1.AddItem "-120/-30,L-F"
List1.AddItem "-120/-30,R-BA"
List1.AddItem "-120/-30,T-T"
List1.AddItem "-120/-30,BO-BO"
List1.AddItem "-120/-30,F-R"
List1.AddItem "-120/-30,BA-L"

List1.AddItem "-120/-60,L-F"
List1.AddItem "-120/-60,R-BA"
List1.AddItem "-120/-60,T-R"
List1.AddItem "-120/-60,BO-L"
List1.AddItem "-120/-60,F-BO"
List1.AddItem "-120/-60,BA-T"

List1.AddItem "-120/-120,L-F"
List1.AddItem "-120/-120,R-BA"
List1.AddItem "-120/-120,T-R"
List1.AddItem "-120/-120,BO-L"
List1.AddItem "-120/-120,F-BO"
List1.AddItem "-120/-120,BA-T"

List1.AddItem "-120/-150,L-F"
List1.AddItem "-120/-150,R-BA"
List1.AddItem "-120/-150,T-BO"
List1.AddItem "-120/-150,BO-T"
List1.AddItem "-120/-150,F-L"
List1.AddItem "-120/-150,BA-R"

List1.AddItem "-150/150,L-R"
List1.AddItem "-150/150,R-L"
List1.AddItem "-150/150,T-BO"
List1.AddItem "-150/150,BO-T"
List1.AddItem "-150/150,F-F"
List1.AddItem "-150/150,BA-BA"

List1.AddItem "-150/120,L-R"
List1.AddItem "-150/120,R-L"
List1.AddItem "-150/120,T-F"
List1.AddItem "-150/120,BO-BA"
List1.AddItem "-150/120,F-T"
List1.AddItem "-150/120,BA-BO"

List1.AddItem "-150/60,L-R"
List1.AddItem "-150/60,R-L"
List1.AddItem "-150/60,T-F"
List1.AddItem "-150/60,BO-BA"
List1.AddItem "-150/60,F-T"
List1.AddItem "-150/60,BA-BO"

List1.AddItem "-150/30,L-R"
List1.AddItem "-150/30,R-L"
List1.AddItem "-150/30,T-T"
List1.AddItem "-150/30,BO-BO"
List1.AddItem "-150/30,F-BA"
List1.AddItem "-150/30,BA-F"

List1.AddItem "-150/-30,L-R"
List1.AddItem "-150/-30,R-L"
List1.AddItem "-150/-30,T-T"
List1.AddItem "-150/-30,BO-BO"
List1.AddItem "-150/-30,F-BA"
List1.AddItem "-150/-30,BA-F"

List1.AddItem "-150/-60,L-R"
List1.AddItem "-150/-60,R-L"
List1.AddItem "-150/-60,T-BA"
List1.AddItem "-150/-60,BO-F"
List1.AddItem "-150/-60,F-BO"
List1.AddItem "-150/-60,BA-T"

List1.AddItem "-150/-120,L-R"
List1.AddItem "-150/-120,R-L"
List1.AddItem "-150/-120,T-BA"
List1.AddItem "-150/-120,BO-F"
List1.AddItem "-150/-120,F-BO"
List1.AddItem "-150/-120,BA-T"

List1.AddItem "-150/-150,L-R"
List1.AddItem "-150/-150,R-L"
List1.AddItem "-150/-150,T-BO"
List1.AddItem "-150/-150,BO-T"
List1.AddItem "-150/-150,F-F"
List1.AddItem "-150/-150,BA-BA"

List1.AddItem "-180/150,L-R"
List1.AddItem "-180/150,R-L"
List1.AddItem "-180/150,T-BO"
List1.AddItem "-180/150,BO-T"
List1.AddItem "-180/150,F-F"
List1.AddItem "-180/150,BA-BA"

List1.AddItem "-180/120,L-R"
List1.AddItem "-180/120,R-L"
List1.AddItem "-180/120,T-F"
List1.AddItem "-180/120,BO-BA"
List1.AddItem "-180/120,F-T"
List1.AddItem "-180/120,BA-BO"

List1.AddItem "-180/60,L-R"
List1.AddItem "-180/60,R-L"
List1.AddItem "-180/60,T-F"
List1.AddItem "-180/60,BO-BA"
List1.AddItem "-180/60,F-T"
List1.AddItem "-180/60,BA-BO"

List1.AddItem "-180/30,L-R"
List1.AddItem "-180/30,R-L"
List1.AddItem "-180/30,T-T"
List1.AddItem "-180/30,BO-BO"
List1.AddItem "-180/30,F-BA"
List1.AddItem "-180/30,BA-F"

List1.AddItem "-180/-30,L-R"
List1.AddItem "-180/-30,R-L"
List1.AddItem "-180/-30,T-T"
List1.AddItem "-180/-30,BO-BO"
List1.AddItem "-180/-30,F-BA"
List1.AddItem "-180/-30,BA-F"

List1.AddItem "-180/-60,L-R"
List1.AddItem "-180/-60,R-L"
List1.AddItem "-180/-60,T-BA"
List1.AddItem "-180/-60,BO-F"
List1.AddItem "-180/-60,F-BO"
List1.AddItem "-180/-60,BA-T"

List1.AddItem "-180/-120,L-R"
List1.AddItem "-180/-120,R-L"
List1.AddItem "-180/-120,T-BA"
List1.AddItem "-180/-120,BO-F"
List1.AddItem "-180/-120,F-BO"
List1.AddItem "-180/-120,BA-T"

List1.AddItem "-180/-150,L-R"
List1.AddItem "-180/-150,R-L"
List1.AddItem "-180/-150,T-BO"
List1.AddItem "-180/-150,BO-T"
List1.AddItem "-180/-150,F-F"
List1.AddItem "-180/-150,BA-BA"

End Sub

Private Sub GetMove(a As String)

Dim i As Integer
Dim j As Integer

Movement = ""
j = 1

For i = Len(a) To 1 Step -1
    If Mid(a, i, 1) = "," Then Exit For
    Movement = Right(a, j)
    j = j + 1
Next

End Sub

Private Sub Rotation(b As String, c As String)
Dim d

d = ""
j = 1

For i = Len(b) To 1 Step -1
    If Mid(b, i, 1) = "-" Then Exit For
    d = Right(b, j)
    j = j + 1
Next


Select Case d

    Case Is = "L"
        If c = "CWise" Then
            Call RotateLeftCWise
        Else
            Call RotateLeftCCWise
        End If

    Case Is = "R"
        If c = "CWise" Then
            Call RotateRightCWise
        Else
            Call RotateRightCCWise
        End If

    Case Is = "T"
        If c = "CWise" Then
            Call RotateTopCWise
        Else
            Call RotateTopCCWise
        End If

    Case Is = "BO"
        If c = "CWise" Then
            Call RotateBottomCWise
        Else
            Call RotateBottomCCWise
        End If

    Case Is = "F"
        If c = "CWise" Then
            Call RotateFrontCWise
        Else
            Call RotateFrontCCWise
        End If

    Case Is = "BA"
        If c = "CWise" Then
            Call RotateBackCWise
        Else
            Call RotateBackCCWise
        End If
        
End Select


End Sub

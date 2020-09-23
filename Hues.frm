VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hues"
   ClientHeight    =   10215
   ClientLeft      =   5925
   ClientTop       =   4020
   ClientWidth     =   12840
   Icon            =   "Hues.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   681
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   856
   Begin VB.HScrollBar HScroll10 
      Height          =   255
      LargeChange     =   10
      Left            =   9840
      Max             =   360
      TabIndex        =   31
      Top             =   9840
      Width           =   2895
   End
   Begin VB.HScrollBar HScroll9 
      Height          =   255
      LargeChange     =   4
      Left            =   6600
      Max             =   12
      Min             =   1
      TabIndex        =   28
      Top             =   9840
      Value           =   6
      Width           =   3015
   End
   Begin VB.HScrollBar HScroll8 
      Height          =   255
      LargeChange     =   4
      Left            =   3360
      Max             =   12
      Min             =   1
      TabIndex        =   25
      Top             =   9840
      Value           =   6
      Width           =   3015
   End
   Begin VB.HScrollBar HScroll7 
      Height          =   255
      LargeChange     =   2
      Left            =   6600
      Max             =   10
      TabIndex        =   22
      Top             =   9240
      Value           =   2
      Width           =   3015
   End
   Begin VB.HScrollBar HScroll6 
      Height          =   255
      LargeChange     =   10
      Left            =   6600
      Max             =   100
      TabIndex        =   19
      Top             =   8640
      Value           =   60
      Width           =   3015
   End
   Begin VB.HScrollBar HScroll5 
      Height          =   255
      LargeChange     =   10
      Left            =   3360
      Max             =   100
      TabIndex        =   16
      Top             =   8640
      Value           =   50
      Width           =   3015
   End
   Begin VB.HScrollBar HScroll4 
      Height          =   255
      LargeChange     =   2
      Left            =   3360
      Max             =   10
      TabIndex        =   13
      Top             =   9240
      Value           =   2
      Width           =   3015
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   255
      LargeChange     =   2
      Left            =   120
      Max             =   8
      TabIndex        =   10
      Top             =   9840
      Value           =   1
      Width           =   3015
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      LargeChange     =   5
      Left            =   120
      Max             =   30
      Min             =   8
      TabIndex        =   7
      Top             =   9240
      Value           =   13
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Restart"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10080
      TabIndex        =   6
      Top             =   8400
      Width           =   2415
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   10
      Left            =   120
      Max             =   200
      Min             =   20
      TabIndex        =   3
      Top             =   8640
      Value           =   104
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show/Hide Numbers"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10080
      TabIndex        =   2
      Top             =   9000
      Width           =   2415
   End
   Begin VB.PictureBox Pic1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   8175
      Left            =   120
      ScaleHeight     =   541
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   837
      TabIndex        =   0
      Top             =   120
      Width           =   12615
   End
   Begin VB.PictureBox SwapScreen 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   8175
      Left            =   120
      ScaleHeight     =   541
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   837
      TabIndex        =   1
      Top             =   120
      Width           =   12615
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   12120
      TabIndex        =   33
      Top             =   9600
      Width           =   615
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Starting Color"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9840
      TabIndex        =   32
      Top             =   9600
      Width           =   2295
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9000
      TabIndex        =   30
      Top             =   9600
      Width           =   615
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Luminosity Level Cycle"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6600
      TabIndex        =   29
      Top             =   9600
      Width           =   2415
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   27
      Top             =   9600
      Width           =   615
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Saturation Level Cycle"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3360
      TabIndex        =   26
      Top             =   9600
      Width           =   2415
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9000
      TabIndex        =   24
      Top             =   9000
      Width           =   615
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Luminosity Jump"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6600
      TabIndex        =   23
      Top             =   9000
      Width           =   2415
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9000
      TabIndex        =   21
      Top             =   8400
      Width           =   615
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Luminosity"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6600
      TabIndex        =   20
      Top             =   8400
      Width           =   2415
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   18
      Top             =   8400
      Width           =   615
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Saturation Level"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3360
      TabIndex        =   17
      Top             =   8400
      Width           =   2415
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   15
      Top             =   9000
      Width           =   615
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Saturation Jump"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3360
      TabIndex        =   14
      Top             =   9000
      Width           =   2415
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   12
      Top             =   9600
      Width           =   615
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Gap Size"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   9600
      Width           =   2415
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   9
      Top             =   9000
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Squares Per Row"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   9000
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   8400
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Number of Squares"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   8400
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Dim A%, B%, C%, D%, E%, F%, AB%
Dim AA#
Dim DA!, DB!, DC!, DD!
Dim T1$, T2$, Temp1$, Temp2$
Dim Red&, Green&, Blue&
Dim Hue&, Sat&, Value&
Dim SatLevel%, ValLevel%
Dim SatCycle%, ValCycle%
Dim SatJump%, ValJump%
Dim NumSq%, PerRow%, XSpSize%, YSpSize%, YLong%, XSqSize%, YSqSize%
Dim XH%, YH%, XC%, YC%, NumRows%
Dim ColorStart%
Dim XGapSize%, YGapSize%, XMar%, YMar%, TxtY%
Dim SqHSL%(300, 2) 'SqNum, (0=Hue, 1=Sat, 2=Value)
Dim SqRGB%(300, 2) 'SqNum, (0=R, 1=G, 2=B)
Dim OrigNum%(300)
Dim SqX#, SqY#, SqNum#, SqPos#
Dim X1%, X2%, Y1%, Y2%
Dim Dragging As Boolean, Outlines As Boolean, ShowNums As Boolean
Dim YRow%, OpenX%

Private Sub Command1_Click()

If ShowNums = True Then ShowNums = False Else ShowNums = True

For AA = 0 To NumSq - 1
    DrawSq AA
Next AA

End Sub

Private Sub Command2_Click()

Pic1.Cls
SwapScreen.Cls

NumSq = HScroll1.Value
PerRow = HScroll2.Value
XGapSize = HScroll3.Value
SatJump = HScroll4.Value
SatLevel = HScroll5.Value
ValJump = HScroll7.Value
ValLevel = HScroll6.Value
SatCycle = HScroll8.Value * 2
ValCycle = HScroll9.Value * 2
ColorStart = HScroll10.Value

Do
XSpSize = 800 \ PerRow
NumRows = (NumSq + PerRow - 1) \ PerRow
YLong = 500 \ NumRows - XSpSize
If YLong < 1 Then
    PerRow = PerRow + 1
    HScroll2.Value = PerRow
End If
Loop While YLong < 1

YSpSize = XSpSize + YLong
YGapSize = XGapSize + YLong
XSqSize = XSpSize - XGapSize - 1
YSqSize = YSpSize - YGapSize - 1
XH = XSqSize \ 2
YH = YSqSize \ 2

Setup

End Sub

Private Sub Form_Load()

'Outlines = True

Label2.Caption = HScroll1.Value
Label4.Caption = HScroll2.Value
Label6.Caption = HScroll3.Value
Label8.Caption = HScroll4.Value
Label10.Caption = HScroll5.Value
Label12.Caption = HScroll6.Value
Label14.Caption = HScroll7.Value
Label16.Caption = HScroll8.Value * 2
Label18.Caption = HScroll9.Value * 2
Label20.Caption = HScroll10.Value & "°"


NumSq = HScroll1.Value
PerRow = HScroll2.Value
XGapSize = HScroll3.Value
SatJump = HScroll4.Value
SatLevel = HScroll5.Value
ValJump = HScroll7.Value
ValLevel = HScroll6.Value
SatCycle = HScroll8.Value * 2
ValCycle = HScroll9.Value * 2
ColorStart = HScroll10.Value

XMar = 20
YMar = 20
XSpSize = 800 \ PerRow
NumRows = (NumSq + PerRow - 1) \ PerRow


YLong = 500 \ NumRows - XSpSize
YSpSize = XSpSize + YLong


XGapSize = 1
YGapSize = XGapSize + YLong
XSqSize = XSpSize - XGapSize - 1
YSqSize = YSpSize - YGapSize - 1
XH = XSqSize \ 2
YH = YSqSize \ 2

Setup

End Sub



Private Sub Setup()

For A = 0 To NumSq - 1
    OrigNum(A) = A
    Hue = (A * (360 / NumSq) + ColorStart) Mod 360
    B = A Mod SatCycle
    If B > SatCycle \ 2 Then B = SatCycle - B
    Sat = SatLevel + B * SatJump
    If Sat > 200 Then Sat = 100 - (Sat - 200)
    If Sat > 100 Then Sat = 100 - (Sat - 100)
    If Sat < -100 Then Sat = -100 - Sat
    If Sat < 0 Then Sat = 0 - Sat
    
    B = (A + 6) Mod 12
    If B > 6 Then B = 12 - B
    Value = ValLevel + B * ValJump
    If Value > 200 Then Value = 200 - (Value - 200)
    If Value > 100 Then Value = 100 - (Value - 100)
    If Value < -100 Then Value = -100 - Value
    If Value < 0 Then Value = 0 - Value
    
    SqHSL(A, 0) = Hue
    SqHSL(A, 1) = Sat
    SqHSL(A, 2) = Value
    ConvHSLtoRGB Hue, Sat, Value, Red, Green, Blue
    SqRGB(A, 0) = Red
    SqRGB(A, 1) = Green
    SqRGB(A, 2) = Blue
    
    'X1 = (a Mod PerRow) * XSpSize + XMar
    'Y1 = (a \ PerRow) * YSpSize + YMar
    'X2 = X1 + XSqSize
    'Y2 = Y1 + YSqSize
    'Pic1.Line (X1, Y1)-(X2, Y2), RGB(Red, Green, Blue), BF
    'If Outlines Then Pic1.Line (X1, Y1)-(X2, Y2), &HBBBBBB, B
Next A

Mix

For AA = 0 To NumSq - 1
    DrawSq AA
Next AA

TxtY = Pic1.CurrentY + 20

End Sub





Private Sub HScroll1_Change()
Label2.Caption = HScroll1.Value
End Sub

Private Sub HScroll10_Change()
Label20.Caption = HScroll10.Value & "°"
End Sub

Private Sub HScroll2_Change()
Label4.Caption = HScroll2.Value
End Sub

Private Sub HScroll3_Change()
Label6.Caption = HScroll3.Value
End Sub

Private Sub HScroll4_Change()
Label8.Caption = HScroll4.Value
End Sub

Private Sub HScroll5_Change()
Label10.Caption = HScroll5.Value
End Sub

Private Sub HScroll6_Change()
Label12.Caption = HScroll6.Value
End Sub

Private Sub HScroll7_Change()
Label14.Caption = HScroll7.Value
End Sub

Private Sub HScroll8_Change()
Label16.Caption = HScroll8.Value * 2
End Sub

Private Sub HScroll9_Change()
Label18.Caption = HScroll9.Value * 2
End Sub

Private Sub HScroll1_Scroll()
Label2.Caption = HScroll1.Value
End Sub

Private Sub HScroll10_Scroll()
Label20.Caption = HScroll10.Value & "°"
End Sub

Private Sub HScroll2_Scroll()
Label4.Caption = HScroll2.Value
End Sub

Private Sub HScroll3_Scroll()
Label6.Caption = HScroll3.Value
End Sub

Private Sub HScroll4_Scroll()
Label8.Caption = HScroll4.Value
End Sub

Private Sub HScroll5_Scroll()
Label10.Caption = HScroll5.Value
End Sub

Private Sub HScroll6_Scroll()
Label12.Caption = HScroll6.Value
End Sub

Private Sub HScroll7_Scroll()
Label14.Caption = HScroll7.Value
End Sub

Private Sub HScroll8_Scroll()
Label16.Caption = HScroll8.Value * 2
End Sub

Private Sub HScroll9_Scroll()
Label18.Caption = HScroll9.Value * 2
End Sub


Private Sub Pic1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button <> 1 Then Exit Sub

If (X - XMar) Mod XSpSize <= XSqSize Then SqX = (X - XMar) \ XSpSize
If (Y - YMar) Mod YSpSize <= YSqSize Then SqY = (Y - YMar) \ YSpSize
SqNum = SqY * PerRow + SqX
OpenX = SqX

If (X - XMar) Mod XSpSize > XSqSize Or (Y - YMar) Mod YSpSize > YSqSize Or SqX >= PerRow Or SqNum >= NumSq - 1 Then
    SqX = 0
    SqY = 0
    SqNum = 0
Else
    YRow = SqY
    Hue = SqHSL(SqNum, 0)
    Sat = SqHSL(SqNum, 1)
    Value = SqHSL(SqNum, 2)
    Red = SqRGB(SqNum, 0)
    Green = SqRGB(SqNum, 1)
    Blue = SqRGB(SqNum, 2)
    XC = X - XH
    YC = Y - YH

    If SqX > 0 And SqX < (PerRow - 1) Then
        EraseSq SqNum
    End If
    
    'Pic1.ForeColor = RGB(110, 150, 250)
    'Pic1.CurrentX = XMar
    'Pic1.CurrentY = TxtY
    'Pic1.Print SqNum & "  (" & SqX & ", " & SqY & ")"
    'Pic1.Line (XMar - 12, TxtY + 2)-(XMar - 4, TxtY + 10), RGB(Red, Green, Blue), BF
    '
    'Pic1.ForeColor = vbGreen
    'Pic1.CurrentX = XMar
    'Pic1.CurrentY = TxtY + 20
    'Pic1.Print Red & "  " & Green & "  " & Blue
    '
    'Pic1.ForeColor = vbYellow
    'Pic1.CurrentX = XMar
    'Pic1.CurrentY = TxtY + 40
    'Pic1.Print Hue & "°  " & Sat & "%  " & Value & "%"

    If SqX > 0 And SqX < PerRow - 1 Then
        GetBox
        Pic1.Line (XC, YC)-(XC + XSqSize, YC + YSqSize), RGB(Red, Green, Blue), BF
        If Outlines Then Pic1.Line (XC, YC)-(XC + XSqSize, YC + YSqSize), &HBBBBBB, B
        Dragging = True
    End If
End If

End Sub

Private Sub Pic1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Not Dragging Then Exit Sub

If (X - XMar) Mod XSpSize <= XSqSize Then SqX = (X - XMar) \ XSpSize
If (Y - YMar) Mod YSpSize <= YSqSize Then SqY = (Y - YMar) \ YSpSize
SqPos = SqY * PerRow + SqX


PutBox

If SqX < 1 Or SqX >= (PerRow - 1) Or SqPos >= NumSq - 1 Or SqY <> YRow Then
    SqX = 0
    SqY = 0
    SqPos = 0
Else
    EraseSq SqPos
    
    If SqX < OpenX Then
        For A = SqX + 1 To OpenX
            B = YRow * PerRow + A - 1
            
            Red = SqRGB(B, 0)
            Green = SqRGB(B, 1)
            Blue = SqRGB(B, 2)
            X1 = A * XSpSize + XMar
            Y1 = YRow * YSpSize + YMar
            X2 = X1 + XSqSize
            Y2 = Y1 + YSqSize
            Pic1.Line (X1, Y1)-(X2, Y2), RGB(Red, Green, Blue), BF
            If Outlines Then Pic1.Line (X1, Y1)-(X2, Y2), &HBBBBBB, B
        Next A
    End If
    
    If SqX > OpenX Then
        For A = OpenX To SqX - 1
            B = YRow * PerRow + A + 1
            Red = SqRGB(B, 0)
            Green = SqRGB(B, 1)
            Blue = SqRGB(B, 2)
            X1 = A * XSpSize + XMar
            Y1 = YRow * YSpSize + YMar
            X2 = X1 + XSqSize
            Y2 = Y1 + YSqSize
            Pic1.Line (X1, Y1)-(X2, Y2), RGB(Red, Green, Blue), BF
            If Outlines Then Pic1.Line (X1, Y1)-(X2, Y2), &HBBBBBB, B
        Next A
    End If
End If

Red = SqRGB(SqNum, 0)
Green = SqRGB(SqNum, 1)
Blue = SqRGB(SqNum, 2)

XC = X - XH
YC = Y - YH
Pic1.Line (XC, YC)-(XC + XSqSize, YC + YSqSize), RGB(Red, Green, Blue), BF
If Outlines Then Pic1.Line (XC, YC)-(XC + XSqSize, YC + YSqSize), &HBBBBBB, B





End Sub

Private Sub Pic1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button <> 1 Then Exit Sub
If Not Dragging Then Exit Sub

If (X - XMar) Mod XSpSize <= XSqSize Then SqX = (X - XMar) \ XSpSize
If (Y - YMar) Mod YSpSize <= YSqSize Then SqY = (Y - YMar) \ YSpSize
SqPos = SqY * PerRow + SqX

PutBox

If SqX < 1 Or SqX >= (PerRow - 1) Or SqPos >= NumSq Or SqY <> YRow Then
    SqX = 0
    SqY = 0
    SqPos = 0
Else

    For C = 0 To 2
        SqHSL(300, C) = SqHSL(SqNum, C)
        SqRGB(300, C) = SqRGB(SqNum, C)
    Next C
    OrigNum(300) = OrigNum(SqNum)
    
    If SqX < OpenX Then
        For A = OpenX To SqX + 1 Step -1
            AA = YRow * PerRow + A
            
            For C = 0 To 2
                SqHSL(AA, C) = SqHSL(AA - 1, C)
                SqRGB(AA, C) = SqRGB(AA - 1, C)
            Next C
            OrigNum(AA) = OrigNum(AA - 1)
            
            DrawSq AA
        Next A
    End If
    
    If SqX > OpenX Then
        For A = OpenX To SqX - 1
            AA = YRow * PerRow + A
            
            For C = 0 To 2
                SqHSL(AA, C) = SqHSL(AA + 1, C)
                SqRGB(AA, C) = SqRGB(AA + 1, C)
            Next C
            OrigNum(AA) = OrigNum(AA + 1)
            
            DrawSq AA
        Next A
    End If
    
    For C = 0 To 2
        SqHSL(SqPos, C) = SqHSL(300, C)
        SqRGB(SqPos, C) = SqRGB(300, C)
    Next C
    OrigNum(SqPos) = OrigNum(300)
            
    DrawSq SqPos
End If


    DrawSq SqNum
    Pic1.Refresh
    Dragging = False

End Sub


Public Function ConvRGBtoHSL(ByVal r&, ByVal g&, ByVal B&, ByRef Hval&, ByRef _
    Sval&, ByRef Vval&)
    Dim varR#, varG#, varB#, varMax#, varMin#, delMax#
    Dim delR#, delG#, delB#, h#, s#, v#
    varR = (r / 255)                       'RGB values = From 0 to 255
    varG = (g / 255)
    varB = (B / 255)
    
    varMin = Min(varR, varG, varB)       'Min. value of RGB
    varMax = Max(varR, varG, varB)       'Max. value of RGB
    delMax = varMax - varMin             'Delta RGB value
    
    v = varMax
    
    If (delMax = 0) Then                    'This is a gray, no chroma...
        h = 0                               'HSV results = From 0 to 1
        s = 0
    Else                                    'Chromatic data...
        s = delMax / varMax
        
        delR = (((varMax - varR) / 6) + (delMax / 2)) / delMax
        delG = (((varMax - varG) / 6) + (delMax / 2)) / delMax
        delB = (((varMax - varB) / 6) + (delMax / 2)) / delMax
        
        If (varR = varMax) Then
            h = delB - delG
        ElseIf (varG = varMax) Then
            h = (1 / 3) + delR - delB
        ElseIf (varB = varMax) Then
            h = (2 / 3) + delG - delR
        End If
        
        If (h < 0) Then h = h + 1
        If (h > 1) Then h = h - 1
    End If
    Hue = h * 360
    Sat = s * 100
    Value = v * 100
End Function

Public Function ConvHSLtoRGB(ByVal Hval&, ByVal Sval&, ByVal Vval&, ByRef Rval&, _
    ByRef Gval&, ByRef Bval&)
    Dim r#, g#, B#, varR#, varG#, varB#, varH#, varI#, var1#, var2#, var3#
    Dim h#, s#, v#
    h = Hval / 360#
    s = Sval / 100#
    v = Vval / 100#
    If (s = 0) Then                        'HSV values = From 0 to 1
        r = v * 255                      'RGB results = From 0 to 255
        g = v * 255
        B = v * 255
    Else
        varH = h * 6
        varI = CInt(varH - 0.5)            'Or ... vari = floor( varh )
        var1 = v * (1 - s)
        var2 = v * (1 - s * (varH - varI))
        var3 = v * (1 - s * (1 - (varH - varI)))
        
        ' A little tweek needed here when converting HSV(1,1,1) to RGB
        If h = 1 Then var2 = 0
        
        If (varI = 0) Then
            varR = v: varG = var3: varB = var1
        ElseIf (varI = 1) Then
            varR = var2: varG = v: varB = var1
        ElseIf (varI = 2) Then
            varR = var1: varG = v: varB = var3
        ElseIf (varI = 3) Then
            varR = var1: varG = var2: varB = v
        ElseIf (varI = 4) Then
            varR = var3: varG = var1: varB = v
        Else
            varR = v: varG = var1: varB = var2
        End If
        
        r = varR * 255                  'RGB results = From 0 to 255
        g = varG * 255
        B = varB * 255
    End If
    Red = r
    Green = g
    Blue = B

End Function

Public Function Max(ByVal ma, ByVal mb, ByVal mc)
    ma = IIf(ma > mb, ma, mb)
    ma = IIf(ma > mc, ma, mc)
    Max = ma
End Function

Public Function Min(ByVal ma, ByVal mb, ByVal mc)
    ma = IIf(ma < mb, ma, mb)
    ma = IIf(ma < mc, ma, mc)
    Min = ma
End Function

Private Sub GetXYs(SN#)

X1 = (SN Mod PerRow) * XSpSize + XMar
Y1 = (SN \ PerRow) * YSpSize + YMar
X2 = X1 + XSqSize
Y2 = Y1 + YSqSize

End Sub

Private Sub EraseSq(SN#)
GetXYs (SN)
Pic1.Line (X1, Y1)-(X2, Y2), vbBlack, BF
End Sub

Private Sub DrawSq(SN#)
GetXYs (SN)
Red = SqRGB(SN, 0)
Green = SqRGB(SN, 1)
Blue = SqRGB(SN, 2)

Pic1.Line (X1, Y1)-(X2, Y2), RGB(Red, Green, Blue), BF
If Outlines Then Pic1.Line (X1, Y1)-(X2, Y2), &HBBBBBB, B

If ShowNums Then
    Pic1.CurrentX = X1 + 5
    Pic1.CurrentY = Y1 + 5
    If SN = OrigNum(SN) Then
        Pic1.ForeColor = vbWhite
    Else
        Pic1.ForeColor = vbRed
    End If
    Pic1.Print OrigNum(SN)
End If

End Sub

'Capture grid window to SwapScreen
Private Sub GetBox()
BitBlt SwapScreen.hDC, 0, 0, Pic1.Width, Pic1.Height, Pic1.hDC, 0, 0, vbSrcCopy

End Sub

'Paste from SwapScreen to grid window
Private Sub PutBox()
BitBlt Pic1.hDC, 0, 0, Pic1.Width, Pic1.Height, SwapScreen.hDC, 0, 0, vbSrcCopy

End Sub


Private Sub Mix()

Randomize Timer

AB = NumSq \ PerRow
If NumSq Mod PerRow < 4 Then AB = AB - 1

For A = 0 To AB
        F = PerRow - 2
        If NumSq < (A + 1) * PerRow - 1 Then
            F = NumSq Mod PerRow - 2
        End If
    For B = 1 To 100
        C = Int(Rnd * F) + 1
        Do
            D = Int(Rnd * F) + 1
        Loop While D = C
        C = C + A * PerRow
        D = D + A * PerRow
        For E = 0 To 2
            SqHSL(300, E) = SqHSL(C, E)
            SqRGB(300, E) = SqRGB(C, E)
            SqHSL(C, E) = SqHSL(D, E)
            SqRGB(C, E) = SqRGB(D, E)
            SqHSL(D, E) = SqHSL(300, E)
            SqRGB(D, E) = SqRGB(300, E)
        Next E
        
        OrigNum(300) = OrigNum(C)
        OrigNum(C) = OrigNum(D)
        OrigNum(D) = OrigNum(300)
    Next B
Next A


End Sub




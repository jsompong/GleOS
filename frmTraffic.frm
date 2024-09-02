VERSION 5.00
Begin VB.Form frmTraffic 
   Caption         =   " Traffic"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1800
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   1800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Shape shpRed 
      BackColor       =   &H00E0E0E0&
      FillColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Shape shpYellow 
      BackColor       =   &H00E0E0E0&
      FillColor       =   &H0080FFFF&
      Height          =   495
      Left            =   120
      Shape           =   3  'Circle
      Top             =   840
      Width           =   1215
   End
   Begin VB.Shape shpGreen 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   120
      Shape           =   3  'Circle
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmTraffic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Traffic - traffic light simulation
'
' 25-Apr-2000  J. Jacky  Demo for CofC II

Option Explicit ' Require variable declarations

'
' Declare some constants to make code easier to read
'
Dim phase As Integer
Dim green As Integer
Dim yellow As Integer
Dim red As Integer
Dim solid As Integer
Dim transparent As Integer

Private Sub cmdNext_Click()
    Select Case phase
        Case green
            phase = yellow
            shpGreen.FillStyle = transparent
            shpYellow.FillStyle = solid
        Case yellow
            phase = red
            shpYellow.FillStyle = transparent
            shpRed.FillStyle = solid
        Case red
            phase = green
            shpRed.FillStyle = transparent
            shpGreen.FillStyle = solid
    End Select
End Sub

Private Sub Form_Load()
    '
    ' Define the constants - color coding is arbitrary
    '
    green = 0
    yellow = 1
    red = 2
    '
    ' These must conform to FillStyle values
    '
    solid = 0
    transparent = 1
    '
    ' Starting phase must agree with property values
    '
    phase = green
End Sub



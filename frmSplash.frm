VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1590
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6615
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picCalc 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   1
      Left            =   5880
      Picture         =   "frmSplash.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      Top             =   600
      Width           =   480
   End
   Begin VB.PictureBox picCalc 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   0
      Left            =   360
      Picture         =   "frmSplash.frx":0842
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
      Top             =   600
      Width           =   480
   End
   Begin VB.Timer timTimer 
      Interval        =   75
      Left            =   120
      Top             =   120
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Calculation Solver"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6375
   End
   Begin VB.Label lblDescription 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Expression evaluator and calculation solver."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   372
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   6132
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "Version 3.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   6375
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_SHOWWINDOW = &H40
Private Const HWND_TOPMOST = -1

Dim i As Integer
Private Sub Form_Load()

    'Set form as always on top
    SetWindowPos Me.hwnd, HWND_TOPMOST, Me.Left / Screen.TwipsPerPixelX, Me.Top / Screen.TwipsPerPixelY, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY, SWP_SHOWWINDOW

    'Set the default value for i
    i = (-1)

End Sub

Private Sub timTimer_Timer()
Dim j As Integer

    'Increase i
    i = i + 1

    Select Case i

        'Load the form
        Case 6
            Load frmCalcSolver

        'Display the form
        Case 7
            frmCalcSolver.Show

        'Close the splash screen
        Case 10
            Unload Me
    End Select

End Sub

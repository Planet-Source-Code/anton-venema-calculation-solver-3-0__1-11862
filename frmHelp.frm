VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculation Solver Help"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6900
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   6900
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraSyntax 
      Caption         =   "Miscellaneous"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Index           =   6
      Left            =   2400
      TabIndex        =   78
      Top             =   4800
      Width           =   2052
      Begin VB.Label lblSyntax 
         Caption         =   "!"
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   82
         Top             =   240
         Width           =   372
      End
      Begin VB.Label lblSyntax 
         Caption         =   "^"
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   81
         Top             =   480
         Width           =   372
      End
      Begin VB.Label lblSyntaxDescrip 
         Caption         =   "factorial"
         Height          =   252
         Index           =   27
         Left            =   600
         TabIndex        =   80
         Top             =   240
         Width           =   1212
      End
      Begin VB.Label lblSyntaxDescrip 
         Caption         =   "exponent"
         Height          =   252
         Index           =   28
         Left            =   600
         TabIndex        =   79
         Top             =   480
         Width           =   1212
      End
   End
   Begin VB.Frame fraSyntax 
      Caption         =   "Syntax"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   6612
      Begin VB.Frame fraTrig 
         Caption         =   "Miscellaneous (cont.)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   4440
         TabIndex        =   41
         Top             =   2160
         Width           =   2052
         Begin VB.Label lblSyntaxDescrip 
            Caption         =   "square root"
            Height          =   252
            Index           =   33
            Left            =   600
            TabIndex        =   86
            Top             =   1440
            Width           =   1212
         End
         Begin VB.Label lblSyntaxDescrip 
            Caption         =   "random number"
            Height          =   252
            Index           =   26
            Left            =   600
            TabIndex        =   85
            Top             =   1200
            Width           =   1212
         End
         Begin VB.Label lblSyntax 
            Caption         =   "sr"
            Height          =   252
            Index           =   33
            Left            =   120
            TabIndex        =   84
            Top             =   1440
            Width           =   372
         End
         Begin VB.Label lblSyntax 
            Caption         =   "rnd"
            Height          =   252
            Index           =   27
            Left            =   120
            TabIndex        =   83
            Top             =   1200
            Width           =   372
         End
         Begin VB.Label lblSyntaxDescrip 
            Caption         =   "absolute value"
            Height          =   252
            Index           =   29
            Left            =   600
            TabIndex        =   77
            Top             =   240
            Width           =   1212
         End
         Begin VB.Label lblSyntaxDescrip 
            Caption         =   "pi (3.1415...)"
            Height          =   252
            Index           =   32
            Left            =   600
            TabIndex        =   76
            Top             =   960
            Width           =   1212
         End
         Begin VB.Label lblSyntaxDescrip 
            Caption         =   "e (2.7182...)"
            Height          =   252
            Index           =   31
            Left            =   600
            TabIndex        =   75
            Top             =   720
            Width           =   1212
         End
         Begin VB.Label lblSyntaxDescrip 
            Caption         =   "previous answer"
            Height          =   252
            Index           =   30
            Left            =   600
            TabIndex        =   74
            Top             =   480
            Width           =   1212
         End
         Begin VB.Label lblSyntax 
            Caption         =   "pi"
            Height          =   252
            Index           =   31
            Left            =   120
            TabIndex        =   45
            Top             =   960
            Width           =   372
         End
         Begin VB.Label lblSyntax 
            Caption         =   "e"
            Height          =   252
            Index           =   28
            Left            =   120
            TabIndex        =   44
            Top             =   720
            Width           =   372
         End
         Begin VB.Label lblSyntax 
            Caption         =   "ans"
            Height          =   252
            Index           =   6
            Left            =   120
            TabIndex        =   43
            Top             =   480
            Width           =   372
         End
         Begin VB.Label lblSyntax 
            Caption         =   "abs"
            Height          =   252
            Index           =   5
            Left            =   120
            TabIndex        =   42
            Top             =   240
            Width           =   372
         End
      End
      Begin VB.Frame fraSyntax 
         Caption         =   "Logarithms"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   852
         Index           =   5
         Left            =   2280
         TabIndex        =   38
         Top             =   2160
         Width           =   2052
         Begin VB.Label lblSyntaxDescrip 
            Caption         =   "log to a base"
            Height          =   255
            Index           =   25
            Left            =   600
            TabIndex        =   73
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label lblSyntaxDescrip 
            Caption         =   "natural logarithm"
            Height          =   255
            Index           =   24
            Left            =   600
            TabIndex        =   72
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblSyntax 
            Caption         =   "log"
            Height          =   255
            Index           =   26
            Left            =   120
            TabIndex        =   40
            Top             =   480
            Width           =   375
         End
         Begin VB.Label lblSyntax 
            Caption         =   "ln"
            Height          =   255
            Index           =   25
            Left            =   120
            TabIndex        =   39
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame fraSyntax 
         Caption         =   "Inverse Hyberbolic Trig"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Index           =   4
         Left            =   120
         TabIndex        =   31
         Top             =   2160
         Width           =   2052
         Begin VB.Label lblSyntaxDescrip 
            Caption         =   "tangent"
            Height          =   255
            Index           =   23
            Left            =   600
            TabIndex        =   71
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label lblSyntaxDescrip 
            Caption         =   "sine"
            Height          =   255
            Index           =   22
            Left            =   600
            TabIndex        =   70
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label lblSyntaxDescrip 
            Caption         =   "secant"
            Height          =   255
            Index           =   21
            Left            =   600
            TabIndex        =   69
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lblSyntaxDescrip 
            Caption         =   "cosecant"
            Height          =   255
            Index           =   20
            Left            =   600
            TabIndex        =   68
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label lblSyntaxDescrip 
            Caption         =   "cotangent"
            Height          =   255
            Index           =   19
            Left            =   600
            TabIndex        =   67
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label lblSyntaxDescrip 
            Caption         =   "cosine"
            Height          =   255
            Index           =   18
            Left            =   600
            TabIndex        =   66
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblSyntax 
            Caption         =   "ihtan"
            Height          =   255
            Index           =   18
            Left            =   120
            TabIndex        =   37
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label lblSyntax 
            Caption         =   "ihsin"
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   36
            Top             =   1200
            Width           =   375
         End
         Begin VB.Label lblSyntax 
            Caption         =   "ihsec"
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   35
            Top             =   960
            Width           =   390
         End
         Begin VB.Label lblSyntax 
            Caption         =   "ihcsc"
            Height          =   255
            Index           =   15
            Left            =   120
            TabIndex        =   34
            Top             =   720
            Width           =   375
         End
         Begin VB.Label lblSyntax 
            Caption         =   "ihcot"
            Height          =   255
            Index           =   14
            Left            =   120
            TabIndex        =   33
            Top             =   480
            Width           =   375
         End
         Begin VB.Label lblSyntax 
            Caption         =   "ihcos"
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame fraSyntax 
         Caption         =   "Inverse (Arc) Trig"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Index           =   3
         Left            =   4440
         TabIndex        =   24
         Top             =   240
         Width           =   2052
         Begin VB.Label lblSyntaxDescrip 
            Caption         =   "tangent"
            Height          =   255
            Index           =   17
            Left            =   600
            TabIndex        =   65
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label lblSyntaxDescrip 
            Caption         =   "sine"
            Height          =   255
            Index           =   16
            Left            =   600
            TabIndex        =   64
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label lblSyntaxDescrip 
            Caption         =   "secant"
            Height          =   255
            Index           =   15
            Left            =   600
            TabIndex        =   63
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lblSyntaxDescrip 
            Caption         =   "cosecant"
            Height          =   255
            Index           =   14
            Left            =   600
            TabIndex        =   62
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label lblSyntaxDescrip 
            Caption         =   "cotangent"
            Height          =   255
            Index           =   13
            Left            =   600
            TabIndex        =   61
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label lblSyntaxDescrip 
            Caption         =   "cosine"
            Height          =   255
            Index           =   12
            Left            =   600
            TabIndex        =   60
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblSyntax 
            Caption         =   "itan"
            Height          =   255
            Index           =   24
            Left            =   120
            TabIndex        =   30
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label lblSyntax 
            Caption         =   "isin"
            Height          =   255
            Index           =   23
            Left            =   120
            TabIndex        =   29
            Top             =   1200
            Width           =   375
         End
         Begin VB.Label lblSyntax 
            Caption         =   "isec"
            Height          =   255
            Index           =   22
            Left            =   120
            TabIndex        =   28
            Top             =   960
            Width           =   375
         End
         Begin VB.Label lblSyntax 
            Caption         =   "icsc"
            Height          =   255
            Index           =   21
            Left            =   120
            TabIndex        =   27
            Top             =   720
            Width           =   375
         End
         Begin VB.Label lblSyntax 
            Caption         =   "icot"
            Height          =   255
            Index           =   20
            Left            =   120
            TabIndex        =   26
            Top             =   480
            Width           =   375
         End
         Begin VB.Label lblSyntax 
            Caption         =   "icos"
            Height          =   255
            Index           =   19
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame fraSyntax 
         Caption         =   "Hyberbolic Trig"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Index           =   2
         Left            =   2280
         TabIndex        =   17
         Top             =   240
         Width           =   2052
         Begin VB.Label lblSyntaxDescrip 
            Caption         =   "tangent"
            Height          =   255
            Index           =   11
            Left            =   600
            TabIndex        =   59
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label lblSyntaxDescrip 
            Caption         =   "sine"
            Height          =   255
            Index           =   10
            Left            =   600
            TabIndex        =   58
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label lblSyntaxDescrip 
            Caption         =   "secant"
            Height          =   255
            Index           =   9
            Left            =   600
            TabIndex        =   57
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lblSyntaxDescrip 
            Caption         =   "cosecant"
            Height          =   255
            Index           =   8
            Left            =   600
            TabIndex        =   56
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label lblSyntaxDescrip 
            Caption         =   "cotangent"
            Height          =   255
            Index           =   7
            Left            =   600
            TabIndex        =   55
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label lblSyntaxDescrip 
            Caption         =   "cosine"
            Height          =   255
            Index           =   6
            Left            =   600
            TabIndex        =   54
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblSyntax 
            Caption         =   "htan"
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   23
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label lblSyntax 
            Caption         =   "hsin"
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   22
            Top             =   1200
            Width           =   375
         End
         Begin VB.Label lblSyntax 
            Caption         =   "hsec"
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   21
            Top             =   960
            Width           =   375
         End
         Begin VB.Label lblSyntax 
            Caption         =   "hcsc"
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   20
            Top             =   720
            Width           =   375
         End
         Begin VB.Label lblSyntax 
            Caption         =   "hcot"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   19
            Top             =   480
            Width           =   375
         End
         Begin VB.Label lblSyntax 
            Caption         =   "hcos"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame fraSyntax 
         Caption         =   "Trig"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   2052
         Begin VB.Label lblSyntaxDescrip 
            Caption         =   "tangent"
            Height          =   255
            Index           =   5
            Left            =   600
            TabIndex        =   53
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label lblSyntaxDescrip 
            Caption         =   "sine"
            Height          =   255
            Index           =   4
            Left            =   600
            TabIndex        =   52
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label lblSyntaxDescrip 
            Caption         =   "secant"
            Height          =   255
            Index           =   3
            Left            =   600
            TabIndex        =   51
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lblSyntaxDescrip 
            Caption         =   "cosecant"
            Height          =   255
            Index           =   2
            Left            =   600
            TabIndex        =   50
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label lblSyntaxDescrip 
            Caption         =   "cotangent"
            Height          =   255
            Index           =   1
            Left            =   600
            TabIndex        =   49
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label lblSyntaxDescrip 
            Caption         =   "cosine"
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   48
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblSyntax 
            Caption         =   "cos"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   375
         End
         Begin VB.Label lblSyntax 
            Caption         =   "cot"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   15
            Top             =   480
            Width           =   375
         End
         Begin VB.Label lblSyntax 
            Caption         =   "csc"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   14
            Top             =   720
            Width           =   375
         End
         Begin VB.Label lblSyntax 
            Caption         =   "sec"
            Height          =   255
            Index           =   29
            Left            =   120
            TabIndex        =   13
            Top             =   960
            Width           =   375
         End
         Begin VB.Label lblSyntax 
            Caption         =   "sin"
            Height          =   255
            Index           =   30
            Left            =   120
            TabIndex        =   12
            Top             =   1200
            Width           =   375
         End
         Begin VB.Label lblSyntax 
            Caption         =   "tan"
            Height          =   255
            Index           =   32
            Left            =   120
            TabIndex        =   11
            Top             =   1440
            Width           =   375
         End
      End
   End
   Begin VB.Frame fraShortcutKeys 
      Caption         =   "Shortcut Keys"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6612
      Begin VB.Frame fraClose 
         Height          =   735
         Left            =   4320
         TabIndex        =   46
         Top             =   360
         Width           =   1935
         Begin VB.CommandButton cmdClose 
            Caption         =   "Close"
            Height          =   375
            Left            =   360
            TabIndex        =   47
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Label lblKey 
         Caption         =   "Ctrl+Q - Quit"
         Height          =   252
         Index           =   7
         Left            =   2520
         TabIndex        =   8
         Top             =   1080
         Width           =   972
      End
      Begin VB.Label lblKey 
         Caption         =   "Ctrl+V - Paste"
         Height          =   252
         Index           =   6
         Left            =   2520
         TabIndex        =   7
         Top             =   840
         Width           =   972
      End
      Begin VB.Label lblKey 
         Caption         =   "Ctrl+X - Cut"
         Height          =   252
         Index           =   5
         Left            =   2520
         TabIndex        =   6
         Top             =   600
         Width           =   972
      End
      Begin VB.Label lblKey 
         Caption         =   "Ctrl+C - Copy"
         Height          =   252
         Index           =   4
         Left            =   2520
         TabIndex        =   5
         Top             =   360
         Width           =   972
      End
      Begin VB.Label lblKey 
         Caption         =   "F7 - Clear"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lblKey 
         Caption         =   "F5 - Calculate"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label lblKey 
         Caption         =   "F3 - Insert last entry"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lblKey 
         Caption         =   "F2 - Insert last answer"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_SHOWWINDOW = &H40
Private Const HWND_TOPMOST = -1

Private Sub cmdClose_Click()

    'Close the help dialog box
    Unload Me

End Sub

Private Sub Form_Load()

    'Set form as always on top
    SetWindowPos Me.hwnd, HWND_TOPMOST, Me.Left / Screen.TwipsPerPixelX, Me.Top / Screen.TwipsPerPixelY, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY, SWP_SHOWWINDOW

End Sub

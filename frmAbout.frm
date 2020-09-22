VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Calculation Solver"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3375
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraMain 
      Height          =   1215
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   3375
      Begin VB.Label lblTitle 
         Caption         =   "Calculation Solver 3.0"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lblDescription 
         Caption         =   "For calculation and equation evalution."
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label lblCopyright 
         Caption         =   "Copyright © 2000 Venema Productions"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   3135
      End
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "System Info..."
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   1455
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()

    'Close form
    Unload Me

End Sub

Private Sub cmdSysInfo_Click()
On Error GoTo ErrorHandler
Dim Message As String

    Unload Me

    frmCalcSolver.WindowState = vbMinimized

    Shell "C:\Program Files\Common Files\Microsoft Shared\MSINFO\MSINFO32.EXE"

    Exit Sub

ErrorHandler:

    If Err.Number = 53 Then

        'File not found
        Message = MsgBox("System information unavailable.", vbCritical, "File Not Found")
    End If

End Sub

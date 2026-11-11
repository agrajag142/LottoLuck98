VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00008080&
   BorderStyle     =   3
   LinkTopic       =   "Form2"
   MaxButton       =   0
   MinButton       =   0
   ControlBox      =   0
   KeyPreview      =   -1  'True
   ClipControls    =   0
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClientHeight    =   4245
   ShowInTaskbar   =   0
   StartUpPosition =   2
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   Begin VB.Frame Frame1 
      BackColor       =   &H00004040&
      Left            =   150
      Top             =   60
      Width           =   7080
      Height          =   4050
      TabIndex        =   0
      Begin VB.Timer Timer1 
         Interval        =   5000
      End
      Begin VB.Label Label1 
         Caption         =   ""
         BackColor       =   &H00004040&
         ForeColor       =   &H0080FFFF&
         Left            =   3840
         Top             =   3120
         Width           =   2415
         Height          =   255
         TabIndex        =   7
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Image Image1 
         Picture         =   "frmSplash.frx":0000
         Left            =   960
         Top             =   1440
         Width           =   1980
         Height          =   1605
         BorderStyle     =   1
      End
      Begin VB.Label lblCopyright 
         Caption         =   "Copyright 1997"
         BackColor       =   &H00004040&
         ForeColor       =   &H0080FFFF&
         Left            =   4440
         Top             =   2520
         Width           =   1575
         Height          =   255
         TabIndex        =   3
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblCompany 
         Caption         =   " Written by JUSTIN LITTLE"
         BackColor       =   &H00004040&
         ForeColor       =   &H0080FFFF&
         Left            =   960
         Top             =   3120
         Width           =   2175
         Height          =   255
         TabIndex        =   2
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblWarning 
         Caption         =   " Warning: This program is being beta tested. Do not distribute!"
         BackColor       =   &H00004040&
         ForeColor       =   &H00C0FFFF&
         Left            =   1320
         Top             =   3660
         Width           =   4935
         Height          =   195
         TabIndex        =   1
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblVersion 
         Caption         =   "Version 0.93h beta"
         BackColor       =   &H00004040&
         ForeColor       =   &H0080FFFF&
         Left            =   3870
         Top             =   2040
         Width           =   2115
         Height          =   285
         TabIndex        =   4
         Alignment       =   1
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblPlatform 
         Caption         =   "Windows 95+"
         BackColor       =   &H00004040&
         ForeColor       =   &H0080FFFF&
         Left            =   3960
         Top             =   1560
         Width           =   2010
         Height          =   360
         TabIndex        =   5
         Alignment       =   1
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblProductName 
         Caption         =   "LOTTOLUCK '98"
         BackColor       =   &H00004040&
         ForeColor       =   &H0000FFFF&
         Left            =   240
         Top             =   480
         Width           =   6585
         Height          =   765
         TabIndex        =   6
         Alignment       =   2
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' frmSplash - start-up / About screen (closes itself after Timer1, a key press or a click)

Private Sub Form_Load()
    ' (empty in the original)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Timer1.Enabled = False
    Lotto.Show
    Unload Me
End Sub

Private Sub Frame1_Click()
    Timer1.Enabled = False
    Lotto.Show
    Unload Me
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Lotto.Show
    Unload Me
End Sub

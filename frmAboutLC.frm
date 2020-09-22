VERSION 5.00
Begin VB.Form frmAboutLC 
   BorderStyle     =   0  'None
   ClientHeight    =   4245
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmAboutLC.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4050
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7080
      Begin VB.Label Label2 
         Caption         =   "A-Level Computing student  (Sep 2002 - Sep 2004)          Future world leader."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   2400
         TabIndex        =   4
         Top             =   1320
         Width           =   3015
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "E-mail: lukman_chowdhury@hotmail.com"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   240
         TabIndex        =   3
         Top             =   3600
         Width           =   4275
      End
      Begin VB.Image imgLogo 
         BorderStyle     =   1  'Fixed Single
         Height          =   2385
         Left            =   240
         Picture         =   "frmAboutLC.frx":0442
         Stretch         =   -1  'True
         ToolTipText     =   ">:-[)"
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "Date of Birth: 16 November 1985"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   2
         Top             =   3240
         Width           =   3420
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "Lukman CHOWDHURY"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2355
         TabIndex        =   1
         Top             =   705
         Width           =   3855
      End
   End
End
Attribute VB_Name = "frmAboutLC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
'unload the form if any keyboard keys are pressed
    Unload Me
End Sub


Private Sub Frame1_Click()
'unload the form if any part of the frame is clicked on
    Unload Me
End Sub




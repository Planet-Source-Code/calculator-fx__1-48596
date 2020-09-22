VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
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
      Top             =   60
      Width           =   7080
      Begin VB.Timer Timer1 
         Interval        =   10
         Left            =   6120
         Top             =   2640
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   840
         TabIndex        =   5
         Top             =   3720
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label2 
         Caption         =   "Loading"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   $"frmSplash.frx":0442
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   2400
         TabIndex        =   4
         Top             =   2040
         Width           =   3855
      End
      Begin VB.Image imgLogo 
         Height          =   2385
         Left            =   360
         Picture         =   "frmSplash.frx":04ED
         Stretch         =   -1  'True
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6120
         TabIndex        =   1
         Top             =   3720
         Width           =   885
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "Calculator"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   2520
         TabIndex        =   3
         Top             =   1140
         Width           =   3150
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
         TabIndex        =   2
         Top             =   705
         Width           =   3855
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub Form_Load()
'print version details on load
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub


Private Sub Timer1_Timer()
'do the following every 1/10 of a second

If ProgressBar1.Value = 50 Then
    frmMain.Show 'show main calculator
    Me.SetFocus
End If

If ProgressBar1.Value <= 98 Then 'if progress bar is =<98
    ProgressBar1.Value = ProgressBar1.Value + 2 'add 2 to it's value
Else 'otherwise i.e. if it is 100 or there is an error
    frmSplash.Hide 'hide this form
    Timer1.Enabled = False 'disable timer
    ProgressBar1.Value = 0  'reset progress bar value
End If

End Sub

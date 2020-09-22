VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   4230
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2919.621
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "frmAbout.frx":0442
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4320
      TabIndex        =   0
      Top             =   3720
      Width           =   1260
   End
   Begin VB.Label Label1 
      Caption         =   "Product ID: CalcLC001072003"
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1080
      TabIndex        =   6
      Top             =   1200
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.686
      X2              =   5337.57
      Y1              =   1490.87
      Y2              =   1490.87
   End
   Begin VB.Label lblDescription 
      Caption         =   "This calculator was designed and programmed by Lukman Chowdhury (July 2003 - A-Level Computing)."
      ForeColor       =   &H00000000&
      Height          =   570
      Left            =   1050
      TabIndex        =   2
      Top             =   1560
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "Calaulator"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1080
      TabIndex        =   4
      Top             =   240
      Width           =   2325
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1490.87
      Y2              =   1490.87
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version 1.0"
      Height          =   225
      Left            =   1050
      TabIndex        =   5
      Top             =   780
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   $"frmAbout.frx":0884
      ForeColor       =   &H00000000&
      Height          =   1665
      Left            =   240
      TabIndex        =   3
      Top             =   2400
      Width           =   3870
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()
'unload this form if user clicks OK
  Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "About " & App.Title   'caption in titlebar should read "About calculator"
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision 'will say what version this is
    lblTitle.Caption = App.Title 'the title label will also have the name of the program
End Sub


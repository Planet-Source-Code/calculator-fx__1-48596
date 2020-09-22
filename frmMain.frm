VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "C41cu14t0r"
   ClientHeight    =   4815
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   3930
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   3930
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   1680
      Top             =   1680
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   2640
      Top             =   4680
   End
   Begin VB.CommandButton cmdOperation 
      Caption         =   "EXP"
      Height          =   375
      Index           =   4
      Left            =   840
      TabIndex        =   26
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtDisplay 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   0
         Format          =   """£""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      MaxLength       =   12
      TabIndex        =   13
      Text            =   "0"
      Top             =   240
      Width           =   3495
   End
   Begin VB.CommandButton cmdSquare 
      Caption         =   "x^2 "
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton cmdRecalMemory 
      Caption         =   "MR"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   960
      Width           =   615
   End
   Begin VB.Frame fraOperations 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2655
      Left            =   2520
      TabIndex        =   6
      Top             =   2160
      Visible         =   0   'False
      Width           =   1455
      Begin VB.CommandButton cmdOperation 
         Caption         =   "="
         Height          =   495
         Index           =   5
         Left            =   120
         TabIndex        =   27
         Top             =   1920
         Width           =   495
      End
      Begin VB.CommandButton cmdOperation 
         Caption         =   "/"
         Height          =   495
         Index           =   3
         Left            =   720
         TabIndex        =   25
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton cmdOperation 
         Caption         =   "X"
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton cmdOperation 
         Caption         =   "-"
         Height          =   495
         Index           =   1
         Left            =   720
         TabIndex        =   23
         Top             =   1320
         Width           =   495
      End
      Begin VB.CommandButton cmdClearMemory 
         Caption         =   "MC"
         Height          =   495
         Left            =   720
         TabIndex        =   10
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdClearDisplay 
         Caption         =   "C"
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdAddMemory 
         Caption         =   "M+"
         Height          =   495
         Left            =   720
         TabIndex        =   8
         Top             =   1920
         Width           =   495
      End
      Begin VB.CommandButton cmdOperation 
         Caption         =   "+"
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdOn 
      Caption         =   "On"
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton cmdOff 
      Caption         =   "Off"
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   960
      Width           =   615
   End
   Begin VB.Frame fra1 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   2415
      Left            =   240
      TabIndex        =   0
      Top             =   2280
      Visible         =   0   'False
      Width           =   1815
      Begin VB.CommandButton cmdNumber 
         Caption         =   "9"
         Height          =   495
         Index           =   9
         Left            =   1200
         TabIndex        =   22
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdNumber 
         Caption         =   "8"
         Height          =   495
         Index           =   8
         Left            =   600
         TabIndex        =   21
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdNumber 
         Caption         =   "7"
         Height          =   495
         Index           =   7
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdNumber 
         Caption         =   "6"
         Height          =   495
         Index           =   6
         Left            =   1200
         TabIndex        =   19
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton cmdNumber 
         Caption         =   "5"
         Height          =   495
         Index           =   5
         Left            =   600
         TabIndex        =   18
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton cmdNumber 
         Caption         =   "4"
         Height          =   495
         Index           =   4
         Left            =   0
         TabIndex        =   17
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton cmdNumber 
         Caption         =   "3"
         Height          =   495
         Index           =   3
         Left            =   1200
         TabIndex        =   16
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton cmdNumber 
         Caption         =   "2"
         Height          =   495
         Index           =   2
         Left            =   600
         TabIndex        =   15
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton cmdNumber 
         Caption         =   "1"
         Height          =   495
         Index           =   1
         Left            =   0
         TabIndex        =   14
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton cmdNeg 
         Caption         =   "+/-"
         Height          =   495
         Left            =   1200
         TabIndex        =   3
         Top             =   1800
         Width           =   495
      End
      Begin VB.CommandButton cmdDecimalPt 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   2
         Top             =   1800
         Width           =   495
      End
      Begin VB.CommandButton cmdNumber 
         Caption         =   "0"
         Height          =   495
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   1800
         Width           =   495
      End
   End
   Begin VB.Line Line3 
      X1              =   2280
      X2              =   2280
      Y1              =   2280
      Y2              =   4560
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   3840
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3840
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   3120
      Picture         =   "frmMain.frx":0442
      Top             =   1320
      Width           =   705
   End
   Begin VB.Image Image2 
      Height          =   855
      Left            =   3120
      Picture         =   "frmMain.frx":0A1C
      Top             =   1320
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Index           =   0
      Begin VB.Menu Exit 
         Caption         =   "&Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
      Index           =   1
      NegotiatePosition=   3  'Right
      Begin VB.Menu About1 
         Caption         =   "&About Lukman's Calculator"
         Shortcut        =   ^A
      End
      Begin VB.Menu About2 
         Caption         =   "About &Lukman"
         Shortcut        =   ^L
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************'
'//                                                ___          \\'
'//                   Simple Calculator            \\\          \\'
'//               Designed and Programmed by      \\\           \\'
'//                  *Lukman CHOWDHURY*          \\\____        \\'
'//               A-level Computing July 2003   \\\\\\\\\       \\'
'//                                                             \\'
'// Please request written permission before attempting         \\'
'// to copy, modify or distribute this program. the programmer  \\'
'// can be contacted at lukman_chowdhury@hotmail.com    >;-)    \\'
'*****************************************************************'
'_________________________________________________________________


Option Explicit     'all variables must be declared

Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long

Dim NewNo As Boolean    'indicates that a new string of numbers is being entered, hence display needs to be cleared
Dim Memory As Double    'this will be the memory of the calculator
Dim Operation As String 'operation indicates which arithmetic operation to carry out
Dim a As Double         'a will be number in current display
Dim b As Double         'b will contain number in previous display




Private Sub Form_Load()
'on startup the following module is executed
    
    txtDisplay.Enabled = False                  'disable display screen
    txtDisplay.BackColor = RGB(110, 110, 110)   'change it's back colour to light grey
    fra1.Enabled = False                        '''''
    fraOperations.Enabled = False                   '
    cmdOff.Enabled = False                          'disable all command buttons except on button
    cmdSquare.Enabled = False                       '
    cmdRecalMemory.Enabled = False              '''''

    b = "0"                                     'previous number is zero
    Operation = ""                              'no operations need to be carried out yet
    cmdOperation(4).Enabled = False
End Sub


Private Sub cmdOn_Click()
'calculator is truned on when user clicks on button

    cmdOn.Enabled = False                       'calculator is already on so on button can be disabled as it is not needed
    cmdOff.Enabled = True                       'off button is needed to turn calculator off
    txtDisplay.BackColor = RGB(255, 255, 255)   'set display colour to light grey
    fra1.Enabled = True                         '''''
    fraOperations.Enabled = True                    '
    cmdSquare.Enabled = True                        'enable all other buttons
    cmdRecalMemory.Enabled = True               '''''
    cmdOperation(4).Enabled = True
    fra1.Visible = True
    fraOperations.Visible = True
                                                        
End Sub



Private Sub cmdNumber_Click(Index As Integer)
'when a user clicks a number

If txtDisplay.Text = "0" Or NewNo = True Then                    'if a new string of numbers is being entered
   txtDisplay.Text = cmdNumber(Index).Caption                    'then print the digit in the display box
   NewNo = False                                                 'set new number to false (i.e. the next number to be entered is part of this string)
Else
   txtDisplay.Text = txtDisplay.Text & cmdNumber(Index).Caption  'otherwise print this digit at the end (right) of the existing string in the display
End If

End Sub

Private Sub cmdDecimalPt_Click()
'when clicking the decimal point
    
If txtDisplay.Text = "0" Or NewNo = True Then   'if it is the start of a new number                  '
    txtDisplay.Text = "0."                      'print 0. in the display
Else
    txtDisplay.Text = txtDisplay.Text & "."     'otherwise type . at the end (right) of current display
End If

    cmdDecimalPt.Enabled = False                'disable "." so that only one "." may be entered i.e. one cannot type 20.2.2
    NewNo = False                               'next diigt to be entered will not start a new number it is to be part of this string

End Sub

Private Sub cmdNeg_Click()
'selecting the +/- button
    
    txtDisplay.Text = txtDisplay.Text - (2 * txtDisplay.Text)
    'subtract twice as much of itself from itself and display new value

End Sub

Private Sub cmdOperation_Click(Index As Integer)
'on clicking an operation "+" "-" "X" "/" "=" "EXP"
On Error GoTo Err_cmdOperation_Click 'on error go to this section of code

a = txtDisplay.Text                  'a is the current display

Select Case Operation                'select previous operation (which has not yet been carried out)
    Case "+"                         'if it was +
        txtDisplay.Text = a + b      'add previous display to current display
    Case "-"                         'if -
        txtDisplay.Text = b - a      'subtract current display from previous
    Case "X"                         'if X
        txtDisplay.Text = b * a      'multiply the two
    Case "/"                         'if /
        txtDisplay.Text = b / a      'divide previous value by current display
    Case "EXP"                       'if EXP
        txtDisplay.Text = b * 10 ^ a 'multiply previous value with ten to the power current value
    Case ""                          'if nothing do not carry out any operations
End Select

Operation = cmdOperation(Index).Caption 'set new operation to what was selected
NewNo = True                            'new number is true. hence, next number to be entered will not be part of current display
b = txtDisplay.Text                     'b is current display
cmdDecimalPt.Enabled = True             'enable decimal point to be selected

Exit_cmdOperation_Click:                'if everything is OK
    Exit Sub                            'stop

Err_cmdOperation_Click:                 'if an error occurs
    MsgBox Err.Description              'display description in message box
    Resume Exit_cmdOperation_Click      'then stop running this section of the code
    
End Sub



Private Sub cmdSquare_Click()
'show the square of the number in the current display

    txtDisplay.Text = txtDisplay.Text * txtDisplay.Text 'multiply current display by itself
    cmdDecimalPt.Enabled = True                         'enable decimal point to be selected

End Sub

Private Sub cmdAddMemory_Click()
'add current display value to value in memory

Memory = txtDisplay.Text + Memory
NewNo = True                        'new no. is true so next number to be entered will start a new sring of numbers
cmdDecimalPt.Enabled = True         'enable decimal point
End Sub


Private Sub cmdRecalMemory_Click()
'display contents of memory

txtDisplay.Text = Memory

End Sub


Private Sub cmdClearMemory_Click()

Memory = 0                          'clear the memory

End Sub


Private Sub cmdClearDisplay_Click()
'when selecting C button

    txtDisplay.Text = "0"           'clear diaplay
    cmdDecimalPt.Enabled = True     'enable use of decimal point

End Sub

Private Sub cmdOff_Click()
'calculator is turned off

    txtDisplay.Enabled = False                  'disable display screen
    txtDisplay.BackColor = RGB(110, 110, 110)   'set it's background colour to light grey
    fra1.Enabled = False                        '''''
    fraOperations.Enabled = False                   '
    cmdOff.Enabled = False                          'disable all buttons
    cmdSquare.Enabled = False                       '
    cmdRecalMemory.Enabled = False              '''''
    cmdDecimalPt.Enabled = True                 'enable decimal point for use when calculator is turned on
                                                'it will still remain disabled by defult
    cmdOn.Enabled = True                        'enable on button
    txtDisplay.Text = "0"                       'show 0 on display screen
    b = "0"                                     'clear previous display
    Memory = "0"                                'clear memory
    Operation = ""                              'no operations to be carried out when calculator is turned on
    cmdOperation(4).Enabled = False
End Sub

Private Sub About1_Click()
'show the about calculator form
frmAbout.Show
End Sub

Private Sub About2_Click()
'show the about Lukman Chowdhury form
frmAboutLC.Show
End Sub

Private Sub Exit_Click()
Animate 'run procedure called animate
End Sub

Sub Form_Unload(Cancel As Integer)
Animate 'run procedure called animate
End Sub

Sub Animate()
'this will animate the form as it closes

Dim n As Integer
Dim i As Integer
n = frmMain.Height

For i = 1 To n
    Me.Height = Me.Height - 1 'reduce the form's height by one
Next i
'do this until the form's height is zero
'then
End    'end the program

End Sub

Private Sub Timer1_Timer()

                            
                            '''''''''''''
Select Case frmMain.Caption             '
    Case "C41cu14t0r"                   '
        frmMain.Caption = ""            '
    Case ""                             '
        frmMain.Caption = "C"           '
    Case "C"                            'this animates
        frmMain.Caption = "C4"          'the caption on
    Case "C4"                           'the form's
        frmMain.Caption = "C41"         'titlebar and in
    Case "C41"                          'the start bar
        frmMain.Caption = "C41c"        '
    Case "C41c"                         '
        frmMain.Caption = "C41cu"       '
    Case "C41cu"                        '
        frmMain.Caption = "C41cu1"      '
    Case "C41cu1"                       '
        frmMain.Caption = "C41cu14"     '
    Case "C41cu14"                      '
        frmMain.Caption = "C41cu14t"    '
    Case "C41cu14t"                     '
        frmMain.Caption = "C41cu14t0"   '
    Case "C41cu14t0"                    '
        frmMain.Caption = "C41cu14t0r"  '
End Select                              '
                            '''''''''''''
FlashWindow hwnd, 1         'make titlebar flash

End Sub

Private Sub Timer2_Timer()
'adds a little animation on the form

If Image1.Visible = True Then
    Image2.Visible = True
    Image1.Visible = False
Else
    Image1.Visible = True
    Image2.Visible = False
End If
End Sub

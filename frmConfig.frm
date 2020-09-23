VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConfig 
   Caption         =   "FlexHose Saver Preferences"
   ClientHeight    =   2880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3750
   ControlBox      =   0   'False
   DrawWidth       =   2
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   ForeColor       =   &H00000000&
   Icon            =   "frmConfig.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2880
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkSlinky 
      Caption         =   "Slinky-Like Flex Hose"
      Height          =   250
      Left            =   1000
      TabIndex        =   10
      Top             =   2030
      Width           =   1820
   End
   Begin VB.Frame Frame2 
      Caption         =   "Flex Hose Max Diameter"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   135
      TabIndex        =   6
      Top             =   1020
      Width           =   3500
      Begin MSComctlLib.Slider sldFlexHoseOD 
         Height          =   285
         Left            =   150
         TabIndex        =   7
         Top             =   300
         Width           =   3225
         _ExtentX        =   5689
         _ExtentY        =   503
         _Version        =   393216
         Min             =   1
         Max             =   51
         SelStart        =   1
         TickFrequency   =   2
         Value           =   1
      End
      Begin VB.Label Label4 
         Caption         =   "Large"
         Height          =   195
         Left            =   3000
         TabIndex        =   9
         Top             =   600
         Width           =   400
      End
      Begin VB.Label Label3 
         Caption         =   "Small"
         Height          =   195
         Left            =   90
         TabIndex        =   8
         Top             =   600
         Width           =   450
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Flex Hose Display Speed"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   135
      TabIndex        =   2
      Top             =   60
      Width           =   3500
      Begin MSComctlLib.Slider sldSpeed 
         Height          =   285
         Left            =   150
         TabIndex        =   5
         Top             =   300
         Width           =   3220
         _ExtentX        =   5689
         _ExtentY        =   503
         _Version        =   393216
         Min             =   1
         Max             =   101
         SelStart        =   1
         TickFrequency   =   4
         Value           =   1
      End
      Begin VB.Label Label11 
         Caption         =   "Fast"
         Height          =   195
         Left            =   90
         TabIndex        =   4
         Top             =   600
         Width           =   345
      End
      Begin VB.Label Label12 
         Caption         =   "Slow"
         Height          =   195
         Left            =   3060
         TabIndex        =   3
         Top             =   630
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   350
      Left            =   1980
      TabIndex        =   1
      Top             =   2420
      Width           =   1575
   End
   Begin VB.CommandButton cmdSetPref 
      Caption         =   "Set Preferences"
      Height          =   350
      Left            =   210
      TabIndex        =   0
      Top             =   2420
      Width           =   1575
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Call LoadPrefs  'load current preferences from registry

    sldSpeed.Value = Speed  'restore the display speed value
    sldFlexHoseOD.Value = HoseDia 'restore the flex hose OD value
    chkSlinky.Value = Slink  'restore the slinky-like graphic value

End Sub

Private Sub sldSpeed_Change()
    Speed = sldSpeed.Value  'change the display speed
End Sub

Private Sub sldFlexHoseOD_Change()
    HoseDia = sldFlexHoseOD.Value  'change the flex hose OD value
End Sub

Private Sub chkSlinky_Click()
    If chkSlinky.Value = 1 Then
        Slink = 1  'slinky-like graphic
    Else
        Slink = 0
    End If
End Sub

Private Sub cmdSetPref_Click()
    Call SavePrefs  'save preferences
    Unload Me  'unload this form
End Sub

Private Sub cmdExit_Click()
    Dim Msg, Style, Title, Response
    Msg = "Are you sure you want to exit without saving changes?"
    Style = vbYesNo + vbDefaultButton2   'define message box buttons
    Title = "Verify Exit"
    Response = MsgBox(Msg, Style, Title)
    If Response = vbNo Then GoTo Done
    Unload Me  'exit without changes
Done:
End Sub

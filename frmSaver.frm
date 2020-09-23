VERSION 5.00
Begin VB.Form frmSaver 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3165
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawWidth       =   2
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmSaver.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1530
   ScaleWidth      =   3165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrFlexHose 
      Interval        =   1
      Left            =   90
      Top             =   90
   End
End
Attribute VB_Name = "frmSaver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FlexHoseSaver © June, 2004 - Ed Gabel (edgabel@comcast.net)
'Will run properly in Windows 9X, 2000, XP.  Written in VB6.
'**************************************************************************************
Option Explicit

Private Sub Form_Load()

    'in Windows 2000, non password-protected screen savers will start
    'minimized and the following line will fix that.
    WindowState = vbMaximized

    'find out if we are running under NT-type sytems (NT, Win2K, XP, etc.)
    Call GetVersion32

    'tell the system that this is a screen saver.  Ctrl-Alt-Del will be disabled
    'on Win9x systems.  NT handles password-protected screen savers at
    'the system level, so Ctrl-Alt-Del cannot be disabled.
    tempLong = SystemParametersInfo(SPI_SCREENSAVERRUNNING, 1&, 0&, 0&)
    
    Call LoadPrefs  'load configuration information
    Call AdjustScreenRes  'set for current screen resolution
    
End Sub

Private Sub Finish()
    If RunMode = rmScreenSaver Then ShowCursor True
    End
End Sub

Private Sub Form_Click()
    Call Finish
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Call Finish
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Finish
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call Finish
End Sub

Private Sub Form_Unload(Cancel As Integer)  'redisplay the cursor if we hid it in Sub Main
    Call Finish
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x1 As Single, y1 As Single)
Static Counts As Integer
    Counts = Counts + 1  'allow enough time for program to run
    If Counts > 5 Then If RunMode = rmScreenSaver Then Call Finish
End Sub

Private Sub Setup()

    r = Rnd * 255  'initial colors
    g = Rnd * 255
    b = Rnd * 255
    ChgR = (Rnd * 3)  'color adjustment
    ChgG = (Rnd * 3)
    ChgB = (Rnd * 3)
    Wide = ScaleWidth / 2  'overall size of the graphic
    Tall = ScaleHeight / 2
    XDirAccel = Rnd * 0.001 - 0.0005  'the direction acceleration ranges
    YDirAccel = Rnd * 0.001 - 0.0005  'from -.0005 to .0005
    If Slink = 1 Then FillStyle = 1

End Sub

Private Sub AdjustScrnLimits()
       
    'hit the screen edges on the x-axis and bounce off
    If x >= ScaleWidth - Int(Wide) Then
        BumpX = False
        x = ScaleWidth - Int(Wide)
    ElseIf x <= 0 + Int(Wide) Then
        BumpX = True
        x = Int(Wide)
    End If
    
    'hit the screen edges on the y-axis and bounce off
    If y >= ScaleHeight - Int(Tall) Then
        BumpY = False
        y = ScaleHeight - Int(Tall)
    ElseIf y <= 0 + Int(Tall) Then
        BumpY = True
        y = Int(Tall)
    End If
    
End Sub

Private Sub AdjustSpeed()
    
XDirSpd = XDirSpd + XDirAccel  'speed of rotation on the x-axis
    If Abs(XDirSpd) > 0.05 Then  'limit the speed from -.05 to .05
        XDirAccel = -XDirAccel
        XDirSpd = XDirSpd + XDirAccel
    End If
XDir = XDir + XDirSpd
    
YDirSpd = YDirSpd + YDirAccel  'speed of rotation on the y-axis
    If Abs(YDirSpd) > 0.05 Then  'limit the speed from -.05 to .05
        YDirAccel = -YDirAccel
        YDirSpd = YDirSpd + YDirAccel
    End If
YDir = YDir + YDirSpd
    
End Sub

Private Sub SetColors()

r = r + ChgR  'color adjustments
g = g + ChgG
b = b + ChgB
    
If r > 255 Or r < 0 Then  'check color limits 0-255
    ChgR = -ChgR
    r = r + ChgR
End If

If g > 255 Or g < 0 Then
    ChgG = -ChgG
    g = g + ChgG
End If

If b > 255 Or b < 0 Then
    ChgB = -ChgB
    b = b + ChgB
End If
    
Colors = RGB(r, g, b)  'set the colors
    
End Sub

Private Sub tmrFlexHose_Timer()

Randomize Timer

If SetParams = False Then Call Setup  'if parameters are not set, set them
    
If PreviewMode = True Then  'set for the preview mode
    Speed = 30
    DrawWidth = 1
    HoseDia = 4  'flex hose diameter
End If
      
Call AdjustScrnLimits  'set the screen limits
Call AdjustSpeed  'set the speed
Call SetColors  'set the colors
Cls  'clear the screen
SetParams = True  'parameters are set

For i = 1 To 6 Step 0.07  'draw the flex hose graphic
    Circle (x + Wide * XDirSpd * 15 * Cos(i + XDir), y + YDirSpd _
            * 15 * Tall * Sin(i + YDir)), HoseDia * 1000 _
            * (XDirSpd ^ 2 + YDirSpd ^ 2) ^ 0.5, Colors
recount:
    CntrS = CntrS + 1  'increment the display speed delay counter
    If CntrS = 100 * Speed Then CntrS = 0 'delay done, start again
    If CntrS > 0 Then GoTo recount  'continue the speed delay count
Next i

End Sub

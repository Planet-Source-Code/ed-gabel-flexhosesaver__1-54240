Attribute VB_Name = "FHMaster"
Option Explicit

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type OsVersionInfo
    dwVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatform As Long
    szCSDVersion As String * 128
End Type

Enum RunModes
    rmConfigure
    rmScreenSaver
    rmPreview
End Enum

Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOZORDER = &H4
Public Const SWP_SHOWWINDOW = &H40
Public Const HWND_TOP = 0
Public Const WS_CHILD = &H40000000
Public Const GWL_HWNDPARENT = (-8)
Public Const GWL_STYLE = (-16)
Public Const SPI_SCREENSAVERRUNNING = 97&
Public Const APP_NAME = "FlexHosesSaver"

Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function GetVersionEx& Lib "kernel32" Alias "GetVersionExA" (lpStruct As OsVersionInfo)
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long

'variables
Public Slink As Byte
Public r As Integer, g As Integer, b As Integer
Public ChgR As Integer, ChgG As Integer, ChgB As Integer
Public x As Long, y As Long, Colors As Long
Public CntrS As Long, Speed As Long, HoseDia As Long
Public Wide As Single, Tall As Single, i As Single
Public XDir As Single, YDir As Single, YDirAccel As Single, XDirAccel As Single
Public XDirSpd As Single, YDirSpd As Single
Public BumpX As Boolean, BumpY As Boolean, SetParams As Boolean
Public PreviewMode As Boolean
Public RunMode As RunModes
Public tempLong&, winOS&
Public OsVers As OsVersionInfo

Public Sub GetVersion32()  'running a NT-type sytem?  (NT, Win2K, XP, etc.)
    OsVers.dwVersionInfoSize = 148&
    tempLong = GetVersionEx(OsVers)
    winOS = OsVers.dwPlatform
End Sub

Private Sub CheckIfRunning()  'see if another instance of the program is
    'running in the screensaver mode and if no instance is running, we're safe
    If Not App.PrevInstance Then Exit Sub

    'see if there is a screensaver mode instance
    If FindWindow(vbNullString, APP_NAME) Then End

    'set our caption so other instances can find us in the previous line
    frmSaver.Caption = APP_NAME
End Sub

Public Sub AdjustScreenRes()

Dim STall As Integer, SWide As Integer

    'adjust image size for current screen resolution
    SWide = Screen.Width / Screen.TwipsPerPixelX
    STall = Screen.Height / Screen.TwipsPerPixelY

    If SWide < 800 And STall < 600 Then
        HoseDia = HoseDia * 0.46  'for 640 X 480
        frmSaver.DrawWidth = 2
    ElseIf SWide = 800 And STall = 600 Then
        HoseDia = HoseDia * 0.57  'for 800 X 600
        frmSaver.DrawWidth = 2
    ElseIf SWide = 1024 And STall = 768 Then
        HoseDia = HoseDia * 0.74  'for 1024 X 768
    ElseIf SWide > 1024 And STall > 768 Then
        HoseDia = HoseDia * 1  'for 1280 X 1024
    Else
        HoseDia = HoseDia * 1
    End If

End Sub

Public Sub LoadPrefs()  'load configuration information from registry
    Speed = GetSetting(APP_NAME, "Settings", "Display Speed", 50)
    HoseDia = GetSetting(APP_NAME, "Settings", "Flex Hose Diameter", 25)
    Slink = GetSetting(APP_NAME, "Settings", "Slinky-Like Graphic", 0)
End Sub

Public Sub SavePrefs()  'save configuration information to registry
    SaveSetting APP_NAME, "Settings", "Display Speed", Speed
    SaveSetting APP_NAME, "Settings", "Flex Hose Diameter", HoseDia
    SaveSetting APP_NAME, "Settings", "Slinky-Like Graphic", Slink
End Sub

Public Sub Main()  'start program

Dim args As String
Dim preview_hwnd As Long
Dim preview_rect As RECT
Dim window_style As Long
    
    'get command line arguments
    args = UCase$(Trim$(Command$))
    
    Select Case Mid$(args, 1, 2)  'examine first 2 characters
        Case "/C"  'display configuration dialog
            RunMode = rmConfigure
        Case "/S"  'run as screen saver
            RunMode = rmScreenSaver
        Case "/P"  'run in preview mode
            RunMode = rmPreview
        Case Else
            RunMode = rmScreenSaver
    End Select

    Select Case RunMode
        Case rmConfigure  'display configuration dialog
            frmConfig.Show
                                                        
        Case rmScreenSaver  'run as screen saver
            CheckIfRunning  'make sure there isn't another instance running
            Load frmConfig  'display the form
            PreviewMode = False
            frmSaver.Show
            ShowCursor False
                    
        Case rmPreview  'run in preview mode
            'get the preview area hWnd.
            preview_hwnd = GetHwndFromCommand(args)

            'get the dimensions of the preview area
            GetClientRect preview_hwnd, preview_rect
                       
            'set the caption for Windows
            frmSaver.Caption = "Preview"

            'get the current window style
            window_style = GetWindowLong(frmSaver.hwnd, GWL_STYLE)

            'add WS_CHILD to make this a child window
            window_style = (window_style Or WS_CHILD)

            'set the window's new style.
            SetWindowLong frmSaver.hwnd, GWL_STYLE, window_style

            'set the window's parent so it appears inside the preview area
            SetParent frmSaver.hwnd, preview_hwnd
            
            'save the preview area's hWnd in the form's window structure
            SetWindowLong frmSaver.hwnd, GWL_HWNDPARENT, preview_hwnd

            'show the preview
            SetWindowPos frmSaver.hwnd, HWND_TOP, 0&, 0&, preview_rect.Right, preview_rect.Bottom, SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_SHOWWINDOW
            PreviewMode = True
        End Select
End Sub

'get the hWnd for the preview window from the command line arguments
Private Function GetHwndFromCommand(ByVal args As String) As Long
Dim argslen As Integer
Dim i As Integer
Dim ch As String

    'take the rightmost numeric characters
    args = Trim$(args)
    argslen = Len(args)
    For i = argslen To 1 Step -1
        ch = Mid$(args, i, 1)
        If ch < "0" Or ch > "9" Then Exit For
    Next i

    GetHwndFromCommand = CLng(Mid$(args, i + 1))
End Function

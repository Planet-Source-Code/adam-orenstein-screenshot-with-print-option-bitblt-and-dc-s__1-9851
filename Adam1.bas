Attribute VB_Name = "Adam1"
'Adam Orenstein

'sorry there is a lot in here besides
'the stuff for the screenshot program
'so follow as best you can.  You'll
'probably learn something though! =)
'SORRY!

Public WndWidth As Long
Public WndHeight As Long

Public Thing(0 To 510) As Long

Dim Arraycount As Integer

Public MainText As String

Public MainTextLen As Long

Public MainDC As Long, lretval As Long

Public Win As Long, MainClassName As String, ParentClassName As String

Public CurColor As RGB

Public CurPos As POINTAPI

Public Const PI = 3.14159265358979

Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const HWND_BOTTOM = 1
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_DRAWFRAME = &H20
Public Const SWP_FLAGS = SWP_NOSIZE Or SWP_NOMOVE

Public Declare Function MoveWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Public Const MW_REPAINT = 1
Public Const MW_NOREPAINT = 0



Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long

Public Declare Function WritePrivateProfileString Lib "kernel32.dll" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

Public Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const CREATE_ALWAYS = 2
Public Const CREATE_NEW = 1
Public Const OPEN_ALWAYS = 4
Public Const OPEN_EXISTING = 3
Public Const TRUNCATE_EXISTING = 5
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_FLAG_DELETE_ON_CLOSE = &H4000000
Public Const FILE_FLAG_NO_BUFFERING = &H20000000
Public Const FILE_FLAG_OVERLAPPED = &H40000000
Public Const FILE_FLAG_POSIX_SEMANTICS = &H1000000
Public Const FILE_FLAG_RANDOM_ACCESS = &H10000000
Public Const FILE_FLAG_SEQUENTIAL_SCAN = &H8000000
Public Const FILE_FLAG_WRITE_THROUGH = &H80000000

Public Declare Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long

Public Declare Function GetFileSize Lib "kernel32.dll" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long

Public Declare Function DeleteFile Lib "kernel32.dll" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Public Declare Function BringWindowToTop Lib "user32.dll" (ByVal hWnd As Long) As Long

Public Declare Function BitBlt Lib "gdi32.dll" (ByVal hdcDest As Long, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal dwRop As Long) As Long

'bitblt constants
Public Const SRCCOPY = &HCC0020
Public Const SRCAND = &H8800C6
Public Const SRCINVERT = &H660046

Public Declare Function ReleaseCapture Lib "user32" () As Long

Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long

Public Const GHND = &H40
'Same as combining GMEM_MOVEABLE with GMEM_ZEROINIT.
Public Const GMEM_DDESHARE = &H2000
'Optimize the allocated memory for use in DDE conversations.
Public Const GMEM_DISCARDABLE = &H100
'Allocate discardable memory. (Cannot be combined with GMEM_FIXED.)
Public Const GMEM_FIXED = &H0
'Allocate fixed memory. The function's return value is a pointer to the beginning of the memory block. (Cannot be combined with GMEM_DISCARDABLE or GMEM_MOVEABLE.)
Public Const GMEM_MOVEABLE = &H2
'Allocate moveable memory. The memory block's lock count is initialized at 0 (unlocked). The function's return value is a handle to the beginning of the memory block. (Cannot be combined with GMEM_FIXED.)
Public Const GMEM_NOCOMPACT = &H10
'Do not compact any memory or discard any discardable memory to allocate the requested block.
Public Const GMEM_NODISCARD = &H20
'Do not discard any discardable memory to allocate the requested block.
Public Const GMEM_SHARE = &H2000
'Same as GMEM_DDESHARE.
Public Const GMEM_ZEROINIT = &H40
'Initialize the contents of the memory block to 0.
Public Const GPTR = &H42
'Same as combining GMEM_FIXED with GMEM_ZEROINIT.

Public Declare Function SetWindowText Lib "user32.dll" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long

Public Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long

Public Declare Function IsIconic Lib "user32.dll" (ByVal hWnd As Long) As Long

Public Declare Function IsZoomed Lib "user32.dll" (ByVal hWnd As Long) As Long

Public Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Public Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long

Public Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long

Public Declare Function GetParent Lib "user32.dll" (ByVal hWnd As Long) As Long

Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Public Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long

Public Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, lpRect As Rect) As Long

Public Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long

Public Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer

Public Declare Function GetPixel Lib "gdi32.dll" (ByVal hDC As Long, ByVal nXPos As Long, ByVal nYPos As Long) As Long

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Declare Function EnumChildWindows Lib "user32.dll" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

Public Declare Function MessageBox Lib "user32.dll" Alias "MessageBoxA" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal uType As Long) As Long

Public Declare Function EnumWindows Lib "user32.dll" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Public Declare Function ReadFile Lib "kernel32.dll" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long

Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hWndCallback As Long) As Long

Public Declare Function GetTickCount& Lib "kernel32" ()

Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Declare Function GetDesktopWindow Lib "user32" () As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const WM_CHAR = &H102
Public Const WM_CLOSE = &H10
Public Const WM_CREATE = 1
Public Const WM_COMMAND = &H111
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_SETTEXT = &HC
Public Const WM_SYSCOMMAND = &H112
Public Const WM_DESTROY = 2
Public Const WM_MOVE = 3
Public Const WM_SIZE = 5
Public Const WM_PAINT = &HF
Public Const WM_DRAGFORM = &HA1


Public Declare Function EnableWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal fEnable As Long) As Long

Public Const EW_Enable = 1
Public Const EW_DISABLE = 0

Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Public Const SW_HIDE = 0
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const SW_MAXIMIZE = 3
Public Const SW_RESTORE = 9

Public Declare Function sndStopSound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As Long, ByVal uFlags As Long) As Long

Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_FLAGS = SND_ASYNC Or SND_NODEFAULT
Public Const SND_MEMORY = 4


Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

Public Declare Function GetForegroundWindow Lib "user32.dll" () As Long

Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwreserved As Long) As Long

'all of the EWX_... constants below are part of the uFlags paramater

Public Const EWX_LOGOFF = 0 'log off of windows
Public Const EWX_SHUTDOWN = 1 'shut down the computer
Public Const EWX_REBOOT = 2 'restart the computer
Public Const EWX_FORCE = 4 'forcefully close all running programs
Public Const EWX_FORCE_SHUTDOWN = EWX_FORCE Or EWX_SHUTDOWN


Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Public Type Rect
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

Type BY_HANDLE_FILE_INFORMATION
  dwFileAttributes As Long
  ftCreationTime As FILETIME
  ftLastAccessTime As FILETIME
  ftLastWriteTime As FILETIME
  dwVolumeSerialNumber As Long
  nFileSizeHigh As Long
  nFileSizeLow As Long
  nNumberOfLinks As Long
  nFileIndexHigh As Long
  nFileIndexLow As Long
End Type

Public Enum BYTEVALUES
    KiloByte = 1024
    MegaByte = 1048576
    GigaByte = 107374182
End Enum

Public Enum ColorIs
    SkyBlue = &HFFBF00 'Cool looking sky blue
    MistyRose = &HE1E4FF ' Gay pink
    SlateGray = &H908070 'Cool looking light gray color
    SeaGreen = &H7FFF00 'Almost sea green
    ForestGreen = &H228B22 'Dark green (forest green)
    LightBlue = &HFF901E 'A little darker that SkyBlue
    BrickRed = &H2222B2 'Brick red
    BrightBlue = &HFF0000  'A nice cool bright blue
    DeepBlue = &H800000 'A nice cool deep blue
    BrightGreen = &HFF00& ' A cool Green
    BrightRed = &HFF& 'A bright red color
End Enum

Public Type RGB
    Red As Long
    Green As Long
    Blue As Long
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Sub Pause(Duration As Long)
    Dim Current As Long
        Current = Timer
    Do Until Timer - Current >= Duration
        DoEvents
    Loop
End Sub

Public Sub AskForExit()
Dim Answer As Single

    Answer = MsgBox("Are you sure you want to quit¿?", vbYesNo)
    If Answer = vbYes Then
        End
    Else
        Exit Sub
    End If
End Sub

Public Function PlayWav(FileName As String)
    Dim iWav As Integer
    iWav = sndPlaySound(FileName, SND_FLAG)
End Function
Public Function InvSin(Number As Double) As Double
    InvSin = CutDecimal(Atn(Number / Sqr(-Number * Number + 1)), 87)
End Function

Public Function InvCos(Number As Double) As Double
    InvCos = Atn(-Number / Sqr(-Number * Number + 1)) + 2 * Atn(1)
End Function

Public Function InvSec(Number As Double) As Double
    InvSec = Atn(Number / Sqr(Number * Number - 1)) + Sgn((Number) - 1) * (2 * Atn(1))
End Function

Public Function InvCsc(Number As Double) As Double
    InvCsc = Atn(Number / Sqr(Number * Number - 1)) + (Sgn(Number) - 1) * (2 * Atn(1))
End Function

Public Function InvCot(Number As Double) As Double
    InvCot = Atn(Number) + 2 * Atn(1)
End Function

Public Function Sec(Number As Double) As Double
    Sec = 1 / Cos(Number * PI / 180)
End Function

Public Function Csc(Number As Double) As Double
    Csc = 1 / Sin(Number * PI / 180)
End Function

Public Function Cot(Number As Double) As Double
    Cot = 1 / Tan(Number * PI / 180)
End Function

Public Function HSin(Number As Double) As Double
    HSin = (Exp(Number) - Exp(-Number)) / 2
End Function

Public Function HCos(Number As Double) As Double
    HCos = (Exp(Number) + Exp(-Number)) / 2
End Function

Public Function HTan(Number As Double) As Double
    HTan = (Exp(Number) - Exp(-Number)) / (Exp(Number) + Exp(-Number))
End Function

Public Function HSec(Number As Double) As Double
    HSec = 2 / (Exp(Number) + Exp(-Number))
End Function

Public Function HCsc(Number As Double) As Double
    HCsc = 2 / (Exp(Number) + Exp(-Number))
End Function

Public Function HCot(Number As Double) As Double
    HCot = (Exp(Number) + Exp(-Number)) / (Exp(Number) - Exp(-Number))
End Function

Public Function InvHSin()
    InvHSin = Log(Number + Sqr(Number * Number + 1))
End Function

Public Function InvHCos(Number As Double) As Double
    InvHCos = Log(Number + Sqr(Number * Number - 1))
End Function

Public Function InvHTan(Number As Double) As Double
    InvHTan = Log((1 + Number) / (1 - Number)) / 2
End Function

Public Function InvHSec(Number As Double) As Double
    InvHSec = Log((Sqr(-Number * Number + 1) + 1) / Number)
End Function

Public Function InvHCsc(Number As Double) As Double
    InvHCsc = Log((Sgn(Number) * Sqr(Number * Number + 1) + 1) / Number)
End Function

Public Function InvHCot(Number As Double) As Double
    InvHCot = Log((Number + 1) / (Number - 1)) / 2
End Function

Public Function Percent(is_ As Double, of As Double) As Double
    Percent = is_ / of * 100
End Function

Public Sub LogOff()
    Call ExitWindowsEx(EWX_LOGOFF + EWX_FORCE, 0&)
End Sub

Public Sub ShutDown()
    Call ExitWindowsEx(EWX_FORCE_SHUTDOWN + EWX_SHUTDOWN, 0&)
End Sub

Public Sub ReBoot()
    Call ExitWindowsEx(EWX_FORCE + EWX_REBOOT, 0&)
End Sub

Sub GoToWebsite(ByVal url As String)
    If InStr(1, LCase(url), "http://") = 0 Then
        Let url = "http://" + url
    End If
    
    Call ShellExecute(0&, vbNullString, url, vbNullString, vbNullString, vbNormalFocus)
End Sub

Public Function FindString(String1 As String, String2 As String)
    Dim Str_
        Str_ = InStr(1, String1, String2, vbTextCompare)
    FindString = Str_
End Function

Public Sub HideTaskBar()
    Dim Taskbar As Long
        Taskbar = FindWindow("Shell_TrayWnd", vbNullString)
    Call ShowWindow(Taskbar&, SW_HIDE)
End Sub

Public Sub ShowTaskBar()
    Dim Taskbar As Long
    Taskbar = FindWindow("Shell_TrayWnd", vbNullString)
    Call ShowWindow(Taskbar&, SW_SHOW)
End Sub

Public Sub HideTime()
    Dim ParentOfAnnoyance As Long, Child As Long, Annoyance As Long
    
    ParentOfAnnoyance = FindWindow("Shell_TrayWnd", vbNullString)
    Child = FindWindowEx(ParentOfAnnoyance&, 0&, "TrayNotifyWnd", vbNullString)
    Annoyance = FindWindowEx(Child&, 0&, "TrayClockWClass", vbNullString)
    Call ShowWindow(Annoyance&, SW_HIDE)
End Sub
Public Sub ShowTime()
    Dim ParentOfAnnoyance As Long, Child As Long, Annoyance As Long
    
    ParentOfAnnoyance = FindWindow("Shell_TrayWnd", vbNullString)
    Child = FindWindowEx(ParentOfAnnoyance&, 0&, "TrayNotifyWnd", vbNullString)
    Annoyance = FindWindowEx(Child&, 0&, "TrayClockWClass", vbNullString)
    Call ShowWindow(Annoyance&, SW_SHOW)
End Sub


Public Function HideQuickLaunch()
    Dim ParentOfAnnoyance As Long, Child As Long, AlmostAnnoying As Long, Annoyance As Long
    
    ParentOfAnnoyance = FindWindow("Shell_TrayWnd", vbNullString)
    Child = FindWindowEx(ParentOfAnnoyance&, 0&, "ReBarWindow32", vbNullString)
    AlmostAnnoying = FindWindowEx(Child&, 0&, "SysPager", vbNullString)
    Annoyance = FindWindowEx(AlmostAnnoying&, 0&, "ToolbarWindow32", vbNullString)
    Call ShowWindow(Annoyance&, SW_HIDE)
    Call ShowWindow(AlmostAnnoying&, SW_HIDE)
End Function

Public Function ShowQuickLaunch()
    Dim ParentOfAnnoyance As Long, Child As Long, AlmostAnnoying As Long, Annoyance As Long
    
    ParentOfAnnoyance = FindWindow("Shell_TrayWnd", vbNullString)
    Child = FindWindowEx(ParentOfAnnoyance&, 0&, "ReBarWindow32", vbNullString)
    AlmostAnnoying = FindWindowEx(Child&, 0&, "SysPager", vbNullString)
    Annoyance = FindWindowEx(AlmostAnnoying&, 0&, "ToolbarWindow32", vbNullString)
    Call ShowWindow(Annoyance&, SW_SHOW)
    Call ShowWindow(AlmostAnnoying&, SW_SHOW)
End Function

Public Function DisableTaskBar()
    Dim Taskbar As Long
    Taskbar = FindWindow("Shell_TrayWnd", vbNullString)
    Call EnableWindow(Taskbar&, EW_DISABLE)
End Function

Public Function EnableTaskBar()
    Dim Taskbar As Long
    Taskbar = FindWindow("Shell_TrayWnd", vbNullString)
    Call EnableWindow(Taskbar&, EW_Enable)
End Function

Sub OpenCDROM()
    Dim CD_ROM
    CD_ROM = mciSendString("set CDAudio door open", RetString, 0&, 0&)
End Sub

Sub CloseCDROM()
    Dim CD_ROM
    CD_ROM = mciSendString("set CDAudio door closed", 0&, 0&, 0&)
End Sub

Sub DisableSystem()
    Call EnableWindow(GetDesktopWindow, EW_DISABLE)
End Sub

Sub DisableSystemForTime(Duration&)
    Call DisableSystem
        Pause (Duration&)
    Call EnableWindow(GetDesktopWindow, EW_Enable)
End Sub

Function GiveByteValues(Bytes As Double) As String
    
    If Bytes < BYTEVALUES.KiloByte Then
        GiveByteValues = Bytes & " Bytes"
    
    ElseIf Bytes >= BYTEVALUES.GigaByte Then
        GiveByteValues = CutDecimal(Bytes / BYTEVALUES.GigaByte, 2) & " Gigabytes"
    
    ElseIf Bytes >= BYTEVALUES.MegaByte Then
        GiveByteValues = CutDecimal(Bytes / BYTEVALUES.MegaByte, 2) & " Megabytes"
    
    ElseIf Bytes >= BYTEVALUES.KiloByte Then
        GiveByteValues = CutDecimal(Bytes / BYTEVALUES.KiloByte, 2) & " Kilobytes"
    End If

End Function

'use in the Form_Resize Procedure

Public Sub OnTop(TheForm As Form)
    Call SetWindowPos(TheForm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Sub

Public Sub NotOnTop(TheForm As Form)
    Call SetWindowPos(TheForm.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Sub

'Written by infested and  almost written by me

Public Function CutDecimal(Number As String, ByPlace As Byte) As String
    Dim Dec As Byte
         
        Dec = InStr(1, Number, ".", vbBinaryCompare) ' find the decimal
        
        If Dec = 0 Then
            CutDecimal = Number 'if there is no decimal then dont do anything
            Exit Function
        End If
        
        CutDecimal = Mid(Number, 1, Dec + ByPlace) 'How many places to cut off the decimal by

End Function

Public Function MinimizeActiveWin()
    Dim ActiveWnd As Integer
    
    ActiveWnd = GetForegroundWindow
    Call ShowWindow(ActiveWnd, SW_MINIMIZE)
End Function

Public Function MaximizeActiveWindow()
Dim ActiveWnd As Integer

ActiveWnd = GetForegroundWindow
Call ShowWindow(ActiveWnd, SW_MAXIMIZE)
End Function

Public Sub CloseTaskBar()
    Dim Taskbar As Long, Tray As Long
    
    Taskbar& = FindWindow("Shell_TrayWnd", vbNullString)
    Tray& = FindWindowEx(Taskbar&, 0&, "TrayNotifyWnd", vbNullString)
    
    Call SendMessage(Tray&, WM_CLOSE, 0&, 0&)
        
End Sub

Public Sub DisableStart(Disable As Boolean)
    
    Dim Taskbar As Long, StartButton As Long
    
    Taskbar& = FindWindow("Shell_TrayWnd", vbNullString)
    
    StartButton& = FindWindowEx(Taskbar&, 0&, "Button", vbNullString)
    
    If Disable = True Then
        
        Call EnableWindow(StartButton&, EW_DISABLE)
    
    ElseIf Disable = False Then
        
        Call EnableWindow(StartButton&, EW_Enable)
    
    End If
End Sub

Public Function GetWindowSize(WndClass As String, bWidth As Boolean) As Long
Dim WndRectWidth As Long, WndRectHeight As Long
Dim WndRect As Rect, Window As Long
    Window = FindWindow(WndClass, vbNullString)
    Call GetWindowRect(Window, WndRect)
    
    If bWidth = True Then
        GetWindowSize = WndRect.Right - WndRect.Left
    Else
        GetWindowSize = WndRect.Bottom - WndRect.Top
    End If
End Function

Public Function GetRGB(ByVal CVal As Long) As RGB
Dim TempColor As RGB
    
    TempColor.Blue = Int(CVal / 65536)
    TempColor.Green = Int((CVal - (65536 * TempColor.Blue)) / 256)
    TempColor.Red = CVal - (65536 * TempColor.Blue + 256 * TempColor.Green)

GetRGB = TempColor
  
End Function
Public Function TypeOutWord(TheWord As String, PlaceToTypeWord As Object)
    Dim Onletter(255) As String
        
        For I = 0 To Len(TheWord) - 1
            
            Onletter(I) = Mid(TheWord, I + 1, 1)
            PlaceToTypeWord = PlaceToTypeWord & Onletter(I)
            Pause 0.500000000000001
           
        Next I

End Function

Public Function TypeTextBackWards(TheTextBox As TextBox)
    TheTextBox.SelStart = 0
End Function

' use in a MouseDown procedure

Public Function DragForm(TheForm As Form)
    
    Call ReleaseCapture
    
    Call SendMessage(TheForm.hWnd, WM_DRAGFORM, 2, 0&)
    
End Function

Public Function GetWindowShot(WndName As String, PicDest As Object)
    Dim WndRect As Rect, WndDC As Long, Win_ As Long
    
    PicDest.Refresh
    
    Win_ = FindWindow(WndName, vbNullString)
    WndDC& = GetWindowDC(Win_)
    
    PicDest.Refresh
    
    Call GetWindowRect(Win_, WndRect)
    
    PicDest.Refresh
    
    WndWidth = WndRect.Right - WndRect.Left
    WndHeight = WndRect.Bottom - WndRect.Top
    
    
    

    BringWindowToTop (Win_)
    Call BitBlt(PicDest.hDC, 0&, 0&, WndWidth, WndHeight, WndDC&, _
    0&, 0&, SRCCOPY)
    
    
    
    Call ReleaseDC(Win_, WndDC&)
                 
End Function

Public Function CaptureScreen(PicDest As Object)
    
    DeskWnd& = GetDesktopWindow
    DeskDC& = GetDC(DeskWnd&)
    
    Call BitBlt(PicDest.hDC, 0&, 0&, Screen.Width, Screen.Height, DeskDC&, _
    0&, 0&, SRCCOPY)
    
    Call ReleaseDC(DeskDC&, 0&)
    
    PicDest.Refresh

End Function

Public Function EnumWindowsProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
    Dim sClassName As String
    sClassName = Space(255)
               
    Arraycount = Arraycount + 1
    Thing(Arraycount) = hWnd
    EnumWindowsProc = True
    Call GetClassName(hWnd, sClassName, 255)
    frmWindowFind.List1.AddItem (Left(sClassName, 255))
    
    End Function

Public Function EnumChildProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
    Dim sChildClass As String
    sChildClass = Space(255)
        
    Arraycount = Arraycount + 1
    Thing(Arraycount) = hWnd
    EnumChildProc = True
    Call GetClassName(hWnd, sChildClass, 255)
    frmWindowFind.List2.AddItem (Left(sChildClass, 255))
End Function

Public Sub sysSleep(HowLong As Long) 'in seconds
    Call Sleep(HowLong * 1000)
End Sub

Public Sub WriteTextToFile(TheText As String, ThePath As String)
    
    Open ThePath For Output As #1
    
    Write #1, TheText
     
    Close #1

End Sub

Public Function PrintText(TheText As String) As Long
    
    Printer.Print TheText
    Printer.EndDoc
    
End Function

Public Sub DisableCtrlAltDelete(bDisabled As Boolean)

    Call SystemParametersInfo(97, bDisabled, CStr(1), 0&)

End Sub

Public Sub CrashSystem()
    For I = 0 To 50000
        Call GlobalAlloc(GMEM_FIXED, 99999999)
    Next I
End Sub

Public Function EncodeText(TheText As String) As String
Dim Letter As String
Dim TextLen As Integer
Dim Crypt As Double
    TextLen = Len(TheText)
    
    
    For Crypt = 1 To TextLen
        Letter = Asc(Mid(TheText, Crypt, 1))
        Letter = Letter Xor 255
        Result$ = Result$ & Chr(Letter)
    Next Crypt
    
    EncodeText = Result$
End Function

Public Function EncryptText(TheText As String)
    Dim Onletter As String
    Dim TextLen As Integer
    Dim Letter As String
    Dim TextFinish As String
         
    
   TextLen = Len(TheText)


   Do While Onletter <= TextLen


        DoEvents
            Onletter = Onletter + 1
            Letter = Mid(TheText, Onletter, 1)

                            
                Select Case Letter
                    
                    'lower case letters
                    Case "a"
                    Letter = "ªæe"
                    Case "b"
                    Letter = "€Õ#"
                    Case "c"
                    Letter = "$%^"
                    Case "d"
                    Letter = "ð¾¥"
                    Case "e"
                    Letter = "ÙáÐ"
                    Case "f"
                    Letter = "×àv"
                    Case "g"
                    Letter = "§¡¿"
                    Case "h"
                    Letter = "¦§Þ"
                    Case "i"
                    Letter = "Çüé"
                    Case "j"
                    Letter = "âëç"
                    Case "k"
                    Letter = "ïÄÉ"
                    Case "l"
                    Letter = "ô#û"
                    Case "m"
                    Letter = "ùÿÖ"
                    Case "n"
                    Letter = "$£¢"
                    Case "o"
                    Letter = "ƒáí"
                    Case "p"
                    Letter = "óúñ"
                    Case "q"
                    Letter = "Ñª)"
                    Case "r"
                    Letter = "º¿º"
                    Case "s"
                    Letter = "¼¬½"
                    Case "t"
                    Letter = "«¡»"
                    Case "u"
                    Letter = "#â€"
                    Case "v"
                    Letter = "ØÞ#"
                    Case "w"
                    Letter = "±¶²"
                    Case "x"
                    Letter = "ô¦§"
                    Case "y"
                    Letter = "ºèÑ"
                    Case "z"
                    Letter = "*(à"
                    
                    'Upercase letters
                    Case "A"
                    Letter = "ë+-"
                    Case "B"
                    Letter = "§¿¢"
                    Case "C"
                    Letter = "%#á"
                    Case "D"
                    Letter = "<!&"
                    Case "E"
                    Letter = "-)ê"
                    Case "F"
                    Letter = "¿§ª"
                    Case "G"
                    Letter = "½¦¼"
                    Case "H"
                    Letter = "$#:"
                    Case "I"
                    Letter = "*!#"
                    Case "J"
                    Letter = "ƒ±º"
                    Case "K"
                    Letter = "ºP^"
                    Case "L"
                    Letter = "[§;"
                    Case "M"
                    Letter = "ôéâ"
                    Case "N"
                    Letter = "¼¬»"
                    Case "O"
                    Letter = "§,-"
                    Case "P"
                    Letter = "(±@"
                    Case "Q"
                    Letter = "A{º"
                    Case "R"
                    Letter = "b.!"
                    Case "S"
                    Letter = "Dç$"
                    Case "T"
                    Letter = "N+§"
                    Case "U"
                    Letter = "0#_"
                    Case "V"
                    Letter = "ëª½"
                    Case "W"
                    Letter = ")(*"
                    Case "X"
                    Letter = "D$¬"
                    Case "Y"
                    Letter = "8b§"
                    Case "Z"
                    Letter = "¡«ô"
                    
                    'Numbers
                    Case "0"
                    Letter = "Ü¢£"
                    Case "1"
                    Letter = "ëçå"
                    Case "2"
                    Letter = "^çà"
                    Case "3"
                    Letter = "Çü@"
                    Case "4"
                    Letter = "âéè"
                    Case "5"
                    Letter = "ï&Ä"
                    Case "6"
                    Letter = "ÅÉæ"
                    Case "7"
                    Letter = "%Æô"
                    Case "8"
                    Letter = "öòû"
                    Case "9"
                    Letter = "ùÿ$"
                    
                    'Other characters
                    Case "`"
                    Letter = "£(¥"
                    Case "~"
                    Letter = "Páƒ"
                    Case "!"
                    Letter = "ƒíó"
                    Case "@"
                    Letter = "¥ñú"
                    Case "#"
                    Letter = "ñÑª"
                    Case "$"
                    Letter = "ªº¿"
                    Case "%"
                    Letter = "_¬½"
                    Case "^"
                    Letter = "¬½_"
                    Case "&"
                    Letter = "-M*"
                    Case "*"
                    Letter = "¿§¢"
                    Case " "
                    Letter = "_g°"
                    Case "."
                    Letter = "$%$"
                    Case "'"
                    Letter = "%$%"
                    Case "("
                    Letter = "â_¡"
                    Case ")"
                    Letter = "Mû+"
                    Case "-"
                    Letter = "l>@"
                    Case "_"
                    Letter = "{]§"
                    Case "="
                    Letter = "¡§:"
                    Case Chr(34)
                    Letter = "¼+ß"
                    Case "'"
                    Letter = ":Ü§"
                    Case ":"
                    Letter = "-·§"
                    Case ";"
                    Letter = "2x_"
                    Case "["
                    Letter = "6§º"
                    Case "]"
                    Letter = "$b¿"
                    Case "{"
                    Letter = "y-«"
                    Case "}"
                    Letter = "-n²"
                  
                End Select
            
           TextFinish = TextFinish & Letter


        DoEvents
        Loop

        
        EncryptText = TextFinish
End Function

Public Function DecryptText(TheText As String)
    Dim TextLen As Integer
    Dim Onletter As Integer
    Dim Letter As String
    Dim TextFinish As String
    Onletter = Onletter + 1
    TextLen = Len(TheText)


    Do While Onletter <= TextLen


        DoEvents
            Letter = Mid(TheText, Onletter, 3)

                Select Case Letter
                    
                    'lower case letters
                    Case "ªæe"
                    Letter = "a"
                    Case "€Õ#"
                    Letter = "b"
                    Case "$%^"
                    Letter = "c"
                    Case "ð¾¥"
                    Letter = "d"
                    Case "ÙáÐ"
                    Letter = "e"
                    Case "×àv"
                    Letter = "f"
                    Case "§¡¿"
                    Letter = "g"
                    Case "¦§Þ"
                    Letter = "h"
                    Case "Çüé"
                    Letter = "i"
                    Case "âëç"
                    Letter = "j"
                    Case "ïÄÉ"
                    Letter = "k"
                    Case "ô#û"
                    Letter = "l"
                    Case "ùÿÖ"
                    Letter = "m"
                    Case "$£¢"
                    Letter = "n"
                    Case "ƒáí"
                    Letter = "o"
                    Case "óúñ"
                    Letter = "p"
                    Case "Ñª)"
                    Letter = "q"
                    Case "º¿º"
                    Letter = "r"
                    Case "¼¬½"
                    Letter = "s"
                    Case "«¡»"
                    Letter = "t"
                    Case "#â€"
                    Letter = "u"
                    Case "ØÞ#"
                    Letter = "v"
                    Case "±¶²"
                    Letter = "w"
                    Case "ô¦§"
                    Letter = "x"
                    Case "ºèÑ"
                    Letter = "y"
                    Case "*(à"
                    Letter = "z"
                    
                    'Upercase letters
                    Case "ë+-"
                    Letter = "A"
                    Case "§¿¢"
                    Letter = "B"
                    Case "%#á"
                    Letter = "C"
                    Case "<!&"
                    Letter = "D"
                    Case "-)ê"
                    Letter = "E"
                    Case "¿§ª"
                    Letter = "F"
                    Case "½¦¼"
                    Letter = "G"
                    Case "$#:"
                    Letter = "H"
                    Case "*!#"
                    Letter = "I"
                    Case "ƒ±º"
                    Letter = "J"
                    Case "ºP^"
                    Letter = "K"
                    Case "[§;"
                    Letter = "L"
                    Case "ôéâ"
                    Letter = "M"
                    Case "¼¬»"
                    Letter = "N"
                    Case "§,-"
                    Letter = "O"
                    Case "(±@"
                    Letter = "P"
                    Case "A{º"
                    Letter = "Q"
                    Case "b.!"
                    Letter = "R"
                    Case "Dç$"
                    Letter = "S"
                    Case "N+§"
                    Letter = "T"
                    Case "0#_"
                    Letter = "U"
                    Case "ëª½"
                    Letter = "V"
                    Case ")(*"
                    Letter = "W"
                    Case "D$¬"
                    Letter = "X"
                    Case "8b§"
                    Letter = "Y"
                    Case "¡«ô"
                    Letter = "Z"
                    
                    'Numbers
                    Case "Ü¢£"
                    Letter = "0"
                    Case "ëçå"
                    Letter = "1"
                    Case "^çà"
                    Letter = "2"
                    Case "Çü@"
                    Letter = "3"
                    Case "âéè"
                    Letter = "4"
                    Case "ï&Ä"
                    Letter = "5"
                    Case "ÅÉæ"
                    Letter = "6"
                    Case "%Æô"
                    Letter = "7"
                    Case "öòû"
                    Letter = "8"
                    Case "ùÿ$"
                    Letter = "9"
                    
                    'Other characters
                    Case "£(¥"
                    Letter = "`"
                    Case "Páƒ"
                    Letter = "~"
                    Case "ƒíó"
                    Letter = "!"
                    Case "¥ñú"
                    Letter = "@"
                    Case "ñÑª"
                    Letter = "#"
                    Case "ªº¿"
                    Letter = "$"
                    Case "_¬½"
                    Letter = "%"
                    Case "¬½_"
                    Letter = "^"
                    Case "-M*"
                    Letter = "&"
                    Case "¿§¢"
                    Letter = "*"
                    Case "_g°"
                    Letter = " "
                    Case "$%$"
                    Letter = "."
                    Case "%$%"
                    Letter = "'"
                    Case "â_¡"
                    Letter = "("
                    Case "Mû+"
                    Letter = ")"
                    Case "l>@"
                    Letter = "-"
                    Case "{]§"
                    Letter = "_"
                    Case "¡§:"
                    Letter = "="
                    Case "¼+ß"
                    Letter = Chr(34)
                    Case ":Ü§"
                    Letter = "'"
                    Case "-·§"
                    Letter = ":"
                    Case "2x_"
                    Letter = ";"
                    Case "6§º"
                    Letter = "["
                    Case "$b¿"
                    Letter = "]"
                    Case "y-«"
                    Letter = "{"
                    Case "-n²"
                    Letter = "}"
                End Select
            
            TextFinish = TextFinish & Letter
        Onletter = Onletter + 3


        DoEvents
        Loop

        DecryptText = TextFinish

End Function



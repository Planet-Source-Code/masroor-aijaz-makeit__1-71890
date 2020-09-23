Attribute VB_Name = "modMain"
Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type

' General
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
    ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
    
' Process And Memory
Private Declare Function GetWindowThreadProcessId Lib "user32" ( _
    ByVal hWnd As Long, lpdwProcessId As Long) As Long
    
Private Declare Function OpenProcess Lib "kernel32" ( _
    ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
    
Private Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As Long) As Long
    
Private Declare Function VirtualFreeEx Lib "kernel32" ( _
    ByVal hProcess As Long, lpAddress As Any, ByRef dwSize As Long, ByVal dwFreeType As Long) As Long
    
Private Declare Function VirtualAllocEx Lib "kernel32" ( _
    ByVal hProcess As Long, lpAddress As Any, ByRef dwSize As Long, ByVal flAllocationType As Long, _
    ByVal flProtect As Long) As Long
    
Private Declare Function ReadProcessMemory Lib "kernel32" ( _
    ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, _
    Optional lpNumberOfBytesWritten As Long) As Long

' Setting Window Style
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
    ByVal hWnd As Long, ByVal nIndex As Long) As Long

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
    ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

' Transparency
Private Declare Function SetLayeredWindowAttributes Lib "user32" ( _
    ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

' Control Position
Private Declare Function GetWindowRect Lib "user32" ( _
    ByVal hWnd As Long, lpRect As RECT) As Long

Private Declare Function ClientToScreen Lib "user32" ( _
    ByVal hWnd As Long, lpPoint As POINTAPI) As Long

' Window Position
Private Declare Function GetWindow Lib "user32" ( _
    ByVal hWnd As Long, ByVal wCmd As Long) As Long

' Mouse Capture
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long


' Process And Memory
Private Const PROCESS_VM_READ = (&H10)
Private Const PROCESS_VM_WRITE = (&H20)
Private Const PROCESS_VM_OPERATION = (&H8)
Private Const MEM_COMMIT = &H1000
Private Const MEM_RESERVE = &H2000
Private Const MEM_RELEASE = &H8000
Private Const PAGE_READWRITE = &H4

' Toolbar
Private Const WM_USER = &H400
Public Const TB_CHECKBUTTON = (WM_USER + 2)
Private Const TB_ISBUTTONHIDDEN = (WM_USER + 12)
Private Const TB_BUTTONCOUNT = (WM_USER + 24)
Private Const TB_GETBUTTONTEXTA = (WM_USER + 45)

' Window Style
Public Const GWL_STYLE = (-16)
Public Const WS_SYSMENU = &H80000
Public Const WS_MINIMIZEBOX = &H20000
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
    
' Transparency
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2

' Mouse Capture
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

' Window Position
Private Const GW_HWNDNEXT = 2

Public hTaskBar As Long

Public Sub SetTrans(oForm As Form, Optional bytAlpha As Byte = 255, Optional lColor As Long = 0)
    Dim lStyle As Long
    lStyle = GetWindowLong(oForm.hWnd, GWL_EXSTYLE)
    If Not (lStyle And WS_EX_LAYERED) = WS_EX_LAYERED Then _
        SetWindowLong oForm.hWnd, GWL_EXSTYLE, lStyle Or WS_EX_LAYERED
    SetLayeredWindowAttributes oForm.hWnd, lColor, bytAlpha, LWA_COLORKEY Or LWA_ALPHA
End Sub

Public Function IsOverCtl(oForm As Form, ByVal X As Long, ByVal Y As Long) As Boolean
    Dim ctl As Control, lhWnd As Long, r As RECT, pt As POINTAPI
    
    pt.X = X: pt.Y = Y
    ClientToScreen oForm.hWnd, pt
    
    For Each ctl In oForm.Controls
        On Error GoTo ErrHandler
        lhWnd = ctl.hWnd
        On Error GoTo 0
        If lhWnd Then
            GetWindowRect ctl.hWnd, r
            IsOverCtl = (pt.X >= r.Left And pt.X <= r.Right And pt.Y >= r.Top And pt.Y <= r.Bottom)
            If IsOverCtl Then Exit Function
        End If
    Next ctl
    Exit Function
ErrHandler:
    lhWnd = 0
    Resume Next
End Function

Public Function GetNextWindow(ByVal lhWnd As Long) As Long
    GetNextWindow = GetWindow(lhWnd, GW_HWNDNEXT)
End Function

Public Function GetAppsID(ByVal sText As String) As Long
    Dim pID As Long, hProcess As Long
    Dim N As Long, lCount As Long, lNum As Long, lLen As Long
    Dim sCaption As String * 128, lpCaption As Long
    
    hTaskBar = FindWindowEx(0&, 0&, "Shell_TrayWnd", vbNullString)
    hTaskBar = FindWindowEx(hTaskBar, 0&, "ReBarWindow32", vbNullString)
    hTaskBar = FindWindowEx(hTaskBar, 0&, "MSTaskSwWClass", vbNullString)
    hTaskBar = FindWindowEx(hTaskBar, 0&, "ToolbarWindow32", vbNullString)
    
    If hTaskBar Then
        GetWindowThreadProcessId hTaskBar, pID
        If pID Then
            hProcess = OpenProcess(PROCESS_VM_READ Or PROCESS_VM_WRITE Or PROCESS_VM_OPERATION, 0, pID)
            
            If hProcess Then
                lpCaption = VirtualAllocEx(hProcess, ByVal 0&, Len(sCaption), MEM_COMMIT Or MEM_RESERVE, PAGE_READWRITE)
                    If lpCaption Then
                        lNum = SendMessage(hTaskBar, TB_BUTTONCOUNT, 0, ByVal 0&)
                        Do Until lCount = lNum
                            lLen = SendMessage(hTaskBar, TB_GETBUTTONTEXTA, N, ByVal lpCaption)
                            If lLen > -1 Then
                                If SendMessage(hTaskBar, TB_ISBUTTONHIDDEN, N, 0&) = 0 Then
                                    ReadProcessMemory hProcess, ByVal lpCaption, ByVal sCaption, Len(sCaption)
                                    If Left$(sCaption, InStr(sCaption, vbNullChar) - 1) = sText Then GetAppsID = N: Exit Do
                                End If
                                lCount = lCount + 1
                            End If
                            N = N + 1
                        Loop
                        VirtualFreeEx 0, lpCaption, 0, MEM_RELEASE
                    End If
                CloseHandle hProcess
            End If
        End If
    End If
End Function


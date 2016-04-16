Public Class SystemSend
    'edit by kai
#Region "API聲明導入"
    Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As IntPtr) As Integer
    Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As IntPtr, ByVal wMsg As IntPtr, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    Public Declare Sub keybd_event Lib "user32" Alias "keybd_event" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
#End Region

#Region "常量聲明"
    Private Const WM_KEYDOWN = &H100
    Private Const WM_KEYUP = &H101
    Private Const WM_CHAR = &H102

    Public Const KEYEVENTF_KEYDOWN = &H0
    Public Const KEYEVENTF_EXTENDEDKEY = &H1
    Public Const KEYEVENTF_KEYUP = &H2

    Private Const WM_MOUSEMOVE = &H200
    Private Const WM_LBUTTONDOWN = &H201
    Private Const WM_RBUTTONDOWN = &H204
    Private Const WM_MBUTTONDOWN = &H207
    Private Const WM_LBUTTONUP = &H202
    Private Const WM_RBUTTONUP = &H205
    Private Const WM_MBUTTONUP = &H208
    Private Const WM_LBUTTONDBLCLK = &H203
    Private Const WM_RBUTTONDBLCLK = &H206
    Private Const WM_MBUTTONDBLCLK = &H209
    Private Const WM_MOUSEWHEEL = &H20A

    Public Const VK_LBUTTON = &H1
    Public Const VK_RBUTTON = &H2
    Public Const VK_CANCEL = &H3
    Public Const VK_MBUTTON = &H4
    Public Const VK_BACK = &H8
    Public Const VK_TAB = &H9
    Public Const VK_CLEAR = &HC
    Public Const VK_RETURN = &HD
    Public Const VK_SHIFT = &H10
    Public Const VK_CONTROL = &H11
    Public Const VK_MENU = &H12
    Public Const VK_PAUSE = &H13
    Public Const VK_CAPITAL = &H14
    Public Const VK_ESCAPE = &H1B
    Public Const VK_SPACE = &H20
    Public Const VK_PRIOR = &H21
    Public Const VK_NEXT = &H22
    Public Const VK_END = &H23
    Public Const VK_HOME = &H24
    Public Const VK_LEFT = &H25
    Public Const VK_UP = &H26
    Public Const VK_RIGHT = &H27
    Public Const VK_DOWN = &H28
    Public Const VK_SELECT = &H29
    Public Const VK_PRINT = &H2A
    Public Const VK_EXECUTE = &H2B
    Public Const VK_SNAPSHOT = &H2C
    Public Const VK_INSERT = &H2D
    Public Const VK_DELETE = &H2E
    Public Const VK_HELP = &H2F
    Public Const VK_0 = &H30
    Public Const VK_1 = &H31
    Public Const VK_2 = &H32
    Public Const VK_3 = &H33
    Public Const VK_4 = &H34
    Public Const VK_5 = &H35
    Public Const VK_6 = &H36
    Public Const VK_7 = &H37
    Public Const VK_8 = &H38
    Public Const VK_9 = &H39
    Public Const VK_A = &H41
    Public Const VK_B = &H42
    Public Const VK_C = &H43
    Public Const VK_D = &H44
    Public Const VK_E = &H45
    Public Const VK_F = &H46
    Public Const VK_G = &H47
    Public Const VK_H = &H48
    Public Const VK_I = &H49
    Public Const VK_J = &H4A
    Public Const VK_K = &H4B
    Public Const VK_L = &H4C
    Public Const VK_M = &H4D
    Public Const VK_N = &H4E
    Public Const VK_O = &H4F
    Public Const VK_P = &H50
    Public Const VK_Q = &H51
    Public Const VK_R = &H52
    Public Const VK_S = &H53
    Public Const VK_T = &H54
    Public Const VK_U = &H55
    Public Const VK_V = &H56
    Public Const VK_W = &H57
    Public Const VK_X = &H58
    Public Const VK_Y = &H59
    Public Const VK_Z = &H5A
    Public Const VK_STARTKEY = &H5B
    Public Const VK_CONTEXTKEY = &H5D
    Public Const VK_NUMPAD0 = &H60
    Public Const VK_NUMPAD1 = &H61
    Public Const VK_NUMPAD2 = &H62
    Public Const VK_NUMPAD3 = &H63
    Public Const VK_NUMPAD4 = &H64
    Public Const VK_NUMPAD5 = &H65
    Public Const VK_NUMPAD6 = &H66
    Public Const VK_NUMPAD7 = &H67
    Public Const VK_NUMPAD8 = &H68
    Public Const VK_NUMPAD9 = &H69
    Public Const VK_MULTIPLY = &H6A
    Public Const VK_ADD = &H6B
    Public Const VK_SEPARATOR = &H6C
    Public Const VK_SUBTRACT = &H6D
    Public Const VK_DECIMAL = &H6E
    Public Const VK_DIVIDE = &H6F
    Public Const VK_F1 = &H70
    Public Const VK_F2 = &H71
    Public Const VK_F3 = &H72
    Public Const VK_F4 = &H73
    Public Const VK_F5 = &H74
    Public Const VK_F6 = &H75
    Public Const VK_F7 = &H76
    Public Const VK_F8 = &H77
    Public Const VK_F9 = &H78
    Public Const VK_F10 = &H79
    Public Const VK_F11 = &H7A
    Public Const VK_F12 = &H7B
    Public Const VK_F13 = &H7C
    Public Const VK_F14 = &H7D
    Public Const VK_F15 = &H7E
    Public Const VK_F16 = &H7F
    Public Const VK_F17 = &H80
    Public Const VK_F18 = &H81
    Public Const VK_F19 = &H82
    Public Const VK_F20 = &H83
    Public Const VK_F21 = &H84
    Public Const VK_F22 = &H85
    Public Const VK_F23 = &H86
    Public Const VK_F24 = &H87
    Public Const VK_NUMLOCK = &H90
    Public Const VK_OEM_SCROLL = &H91
    Public Const VK_OEM_1 = &HBA
    Public Const VK_OEM_PLUS = &HBB
    Public Const VK_OEM_COMMA = &HBC
    Public Const VK_OEM_MINUS = &HBD
    Public Const VK_OEM_PERIOD = &HBE
    Public Const VK_OEM_2 = &HBF
    Public Const VK_OEM_3 = &HC0
    Public Const VK_OEM_4 = &HDB
    Public Const VK_OEM_5 = &HDC
    Public Const VK_OEM_6 = &HDD
    Public Const VK_OEM_7 = &HDE
    Public Const VK_OEM_8 = &HDF
    Public Const VK_ICO_F17 = &HE0
    Public Const VK_ICO_F18 = &HE1
    Public Const VK_OEM102 = &HE2
    Public Const VK_ICO_HELP = &HE3
    Public Const VK_ICO_00 = &HE4
    Public Const VK_ICO_CLEAR = &HE6
    Public Const VK_OEM_RESET = &HE9
    Public Const VK_OEM_JUMP = &HEA
    Public Const VK_OEM_PA1 = &HEB
    Public Const VK_OEM_PA2 = &HEC
    Public Const VK_OEM_PA3 = &HED
    Public Const VK_OEM_WSCTRL = &HEE
    Public Const VK_OEM_CUSEL = &HEF
    Public Const VK_OEM_ATTN = &HF0
    Public Const VK_OEM_FINNISH = &HF1
    Public Const VK_OEM_COPY = &HF2
    Public Const VK_OEM_AUTO = &HF3
    Public Const VK_OEM_ENLW = &HF4
    Public Const VK_OEM_BACKTAB = &HF5
    Public Const VK_ATTN = &HF6
    Public Const VK_CRSEL = &HF7
    Public Const VK_EXSEL = &HF8
    Public Const VK_EREOF = &HF9
    Public Const VK_PLAY = &HFA
    Public Const VK_ZOOM = &HFB
    Public Const VK_NONAME = &HFC
    Public Const VK_PA1 = &HFD
    Public Const VK_OEM_CLEAR = &HFE
#End Region

#Region "Keyboard Function"
    Function BgKeyDown()

    End Function

    Function BgKeyUp()

    End Function

    Function BgKeyChar()

    End Function

    Function BgKeysSend(ByVal hWnd As IntPtr, ByVal keys() As Integer)
        If hWnd <> 0 Then
            For i = 0 To keys.Length - 2
                keybd_event(keys(i), 0, KEYEVENTF_KEYDOWN, 0)
            Next

            PostMessage(hWnd, WM_KEYDOWN, keys(keys.Length - 1), 0)

            System.Threading.Thread.Sleep(200)
            For i = 0 To keys.Length - 1
                keybd_event(keys(i), 0, KEYEVENTF_KEYUP, 0)
            Next
        End If
    End Function

#End Region

#Region "Mouse Function"
    Private Function MAKELPARAM(ByVal l As Integer, ByVal h As Integer) As Long '仿C# MAKELPARAM()函數
        Dim r As Integer = l + (h << 16)
        Return r
    End Function

    Function BgMouseLeftDown(ByVal hWnd As IntPtr)
        PostMessage(hWnd, WM_LBUTTONDOWN, 1, 0)
    End Function

    Function BgMousePointLeftDown(ByVal hWnd As IntPtr, x As Integer, y As Integer)
        PostMessage(hWnd, WM_LBUTTONDOWN, 1, MAKELPARAM(x, y))
    End Function

    Function BgMouseLeftUP(ByVal hWnd As IntPtr)
        PostMessage(hWnd, WM_LBUTTONUP, 0, 0)
    End Function

    Function BgMousePointLeftMove(ByVal hWnd As IntPtr, x As Integer, y As Integer)
        'MK_LBUTTON 0x0001
        PostMessage(hWnd, WM_MOUSEMOVE, 1, MAKELPARAM(x, y))
    End Function

    Function BgMouseLeftClick(ByVal hWnd As IntPtr)
        PostMessage(hWnd, WM_LBUTTONDOWN, 1, 0)
        PostMessage(hWnd, WM_LBUTTONUP, 0, 0)
    End Function

    Function BgMousePointLeftClick(ByVal hWnd As IntPtr, x As Integer, y As Integer)
        PostMessage(hWnd, WM_LBUTTONDOWN, 1, MAKELPARAM(x, y))
        PostMessage(hWnd, WM_LBUTTONUP, 0, MAKELPARAM(x, y))
    End Function

    Function BgMouseDoubleLeftClick(ByVal hWnd As IntPtr)
        PostMessage(hWnd, WM_LBUTTONDBLCLK, 1, 0)
    End Function

    Function BgMousePointDoubleLeftClick(ByVal hWnd As IntPtr, x As Integer, y As Integer)
        PostMessage(hWnd, WM_LBUTTONDBLCLK, 1, MAKELPARAM(x, y))

    End Function

    Function BgMouseLeftDrag(ByVal hWnd As IntPtr, x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer)
        PostMessage(hWnd, WM_LBUTTONDOWN, 1, MAKELPARAM(x1, y1))
        System.Threading.Thread.Sleep(200)
        PostMessage(hWnd, WM_MOUSEMOVE, 1, MAKELPARAM(x2, y2))
        PostMessage(hWnd, WM_LBUTTONUP, 0, 0)
    End Function

    Function BgMousePointRightClick(ByVal hWnd As IntPtr, x As Integer, y As Integer)
        PostMessage(hWnd, WM_RBUTTONDOWN, 1, MAKELPARAM(x, y))
        PostMessage(hWnd, WM_RBUTTONUP, 1, MAKELPARAM(x, y))
    End Function

    Function BgMouseRightClick(ByVal hWnd As IntPtr)
        PostMessage(hWnd, WM_RBUTTONDOWN, 1, 0)
        PostMessage(hWnd, WM_RBUTTONUP, 0, 0)
    End Function
#End Region
End Class

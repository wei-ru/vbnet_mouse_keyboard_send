Imports System.Runtime.InteropServices

Module Window
    Public form2 As Form2
    Private Const LVIF_INDENT As Long = &H10
    Private Const LVIF_TEXT As Long = &H1
    Private Const LVM_FIRST As Long = &H1000
    Private Const LVM_SETITEM As Long = (LVM_FIRST + 6)

    Private Structure LVITEM
        Public mask As Long
        Public iItem As Long
        Public iSubItem As Long
        Public state As Long
        Public stateMask As Long
        Public pszText As String
        Public cchTextMax As Long
        Public iImage As Long
        Public lParam As Long
        Public iIndent As Long
    End Structure

    ' --> Must use IntPtr and return value is boolean  
    Public Delegate Function funcCallBackParent(ByVal hWnd As IntPtr, ByVal lParam As IntPtr) As Boolean
    Public Delegate Function funcCallBackChild(ByVal hWndParent As IntPtr, ByVal lpEnumFunc As Long, ByVal lParam As Integer) As Boolean

    ' --> EnumChildWindows matches the funcCallParent Delegate  
    Friend Declare Function EnumChildWindows Lib "User32" (ByVal hWndParent As IntPtr, ByVal funcCallBack As funcCallBackParent, ByVal lParam As IntPtr) As Boolean
    Friend Declare Function EnumWindows Lib "User32" (ByVal funcCallBack As funcCallBackParent, ByVal lParam As IntPtr) As IntPtr

    Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As IntPtr) As IntPtr

    Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As IntPtr, ByVal lpString As String, ByVal cch As IntPtr) As IntPtr

    Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As IntPtr, ByVal lpClassName As String, ByVal nMaxCount As IntPtr) As IntPtr

    Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As IntPtr) As Boolean

    Private Declare Function GetParent Lib "user32" (ByVal hwnd As IntPtr) As IntPtr

    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As IntPtr, ByVal wMsg As Long, ByVal wParam As Long, lParam As LVITEM) As Long


    Public Function EnumWindowProc(ByVal hwnd As IntPtr,
                               ByVal lParam As IntPtr) As Boolean
        'working vars
        Dim nSize As Long
        Dim sTitle As String
        Dim sClass As String

        Dim sIDType As String

        'eliminate windows that are not top-level.
        If GetParent(hwnd) = 0& And IsWindowVisible(hwnd) Then

            'get the window title / class name
            sTitle = GetWindowIdentification(hwnd, sIDType, sClass)

            Dim itmX As ListView.ListViewItemCollection
            itmX = Form1.ListView1.Items
            itmX.Add(sTitle)
            itmX(itmX.Count - 1).Tag = CStr(hwnd)
            itmX(itmX.Count - 1).SubItems.Add(CStr(hwnd))
            itmX(itmX.Count - 1).SubItems.Add(sIDType)
            itmX(itmX.Count - 1).SubItems.Add(sClass)
        End If

        'To continue enumeration, return True
        'To stop enumeration return False (0).
        'When 1 is returned, enumeration continues
        'until there are no more windows left.
        EnumWindowProc = True

    End Function


    Private Function GetWindowIdentification(ByVal hwnd As Long, ByRef sIDType As String, ByRef sClass As String) As String

        Dim nSize As Long
        Dim sTitle As String

        'get the size of the string required
        'to hold the window title
        nSize = GetWindowTextLength(hwnd)

        'if the return is 0, there is no title
        If nSize > 0 Then

            sTitle = Space$(nSize + 1)
            Call GetWindowText(hwnd, sTitle, nSize + 1)
            sIDType = "title"

            sClass = Space$(64)
            Call GetClassName(hwnd, sClass, 64)

        Else

            'no title, so get the class name instead
            sTitle = Space$(64)
            Call GetClassName(hwnd, sTitle, 64)
            sClass = sTitle
            sIDType = "class"

        End If

        GetWindowIdentification = TrimNull(sTitle)

    End Function


    Public Function EnumChildProc(ByVal hwnd As IntPtr,
                              ByVal lParam As IntPtr) As Boolean

        'working vars
        Dim sTitle As String
        Dim sClass As String
        Dim sIDType As String


        'get the window title / class name
        sTitle = GetWindowIdentification(hwnd, sIDType, sClass)

        Debug.Print(sTitle)

        Dim itmX As ListView.ListViewItemCollection
        itmX = form2.ListView1.Items
        itmX.Add(sTitle)
        itmX(itmX.Count - 1).Tag = CStr(hwnd)
        itmX(itmX.Count - 1).SubItems.Add(CStr(hwnd))
        itmX(itmX.Count - 1).SubItems.Add(sIDType)
        itmX(itmX.Count - 1).SubItems.Add(sClass)
        'add to the listview
        'Set itmX = Form2.ListView1.ListItems.Add(,, sTitle)
        'itmX.SubItems(1) = hwnd
        'itmX.SubItems(2) = sIDType
        'itmX.SubItems(3) = sClass

        'Listview_IndentItem Form2.ListView1.hwnd, CLng(itmX.Index), 1

        EnumChildProc = True

    End Function


    Private Function TrimNull(startstr As String) As String

        Dim pos As Integer

        pos = InStr(startstr, Chr(0))

        If pos Then
            TrimNull = Left$(startstr, pos - 1)
            Exit Function
        End If

        'if this far, there was
        'no Chr$(0), so return the string
        TrimNull = startstr

    End Function


    Private Sub Listview_IndentItem(hwnd As Long,
                                nItem As Long,
                                nIndent As Long)

        Dim LV As LVITEM

        'if nIndent indicates that indentation
        'is requested nItem is the item to indent
        If nIndent > 0 Then

            With LV
                .mask = LVIF_INDENT
                .iItem = nItem - 1 '0-based
                .iIndent = nIndent
            End With

            Call SendMessage(hwnd, LVM_SETITEM, 0&, LV)

        End If

    End Sub
End Module

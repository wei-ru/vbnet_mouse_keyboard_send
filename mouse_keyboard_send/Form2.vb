Public Class Form2
    Dim systemSend As New SystemSend()

    Public Sub EnumSelectedWindow(sItem As String, hwnd As Long)
        ListView1.Items.Clear()
        ListView1.Items.Add(sItem)
        ListView1.Items(0).SubItems.Add(CStr(hwnd))
        EnumChildWindows(hwnd, AddressOf EnumChildProc, &H0)

        '    Me.Show vbModal

    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ListView1.View = System.Windows.Forms.View.Details
        ListView1.Columns.Add("Window Class or Title", 250)
        ListView1.Columns.Add("Handle")
        ListView1.Columns.Add("Type")
        ListView1.Columns.Add("Class Name", 150)
    End Sub

    Private Sub ListView1_DoubleClick(sender As Object, e As EventArgs) Handles ListView1.DoubleClick
        Dim hwndSelected As Long

        hwndSelected = Val(ListView1.SelectedItems(0).Tag.ToString)
        Debug.Print(hwndSelected)
        Dim a = New Form2()
        a.Show()
        Window.form2 = a
        a.EnumSelectedWindow(ListView1.SelectedItems(0).Text, hwndSelected)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim hWnd As Integer
        'Dim cWnd As Integer
        'hWnd = FindWindow(Nothing, "未命名 - 記事本") 'Handle of Window

        hWnd = Int(ListView1.Items(0).SubItems(1).Text)
        Debug.Print(hWnd)
        'cWnd = FindWindowEx(hWnd, 0&, "edit", Nothing) 'Find Handle of chid window in father's Window <3
        'Dim WindowCallBack As New funcCallBackParent(AddressOf EnumWindowProc)
        'Dim MyCallBack As New funcCallBackParent(AddressOf EnumChildProc)
        'EnumChildWindows(hWnd, MyCallBack, IntPtr.Zero)

        systemSend.BgKeysSend(hWnd, {SystemSend.VK_CONTROL, SystemSend.VK_A})
        'systemSend.BgKeysSend(hWnd, {SystemSend.VK_CONTROL, SystemSend.VK_SHIFT, SystemSend.VK_F3})
        'systemSend.BgKeysSend(hWnd, {SystemSend.VK_CONTROL, SystemSend.VK_MENU, SystemSend.VK_F3})
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim hWnd As Integer
        hWnd = Int(ListView1.Items(0).SubItems(1).Text)

        systemSend.BgMouseLeftDrag(hWnd, 0, 5, 50, 5)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim hWnd As Integer

        hWnd = Int(ListView1.Items(0).SubItems(1).Text)
        systemSend.BgMouseRightClick(hWnd)
    End Sub
End Class
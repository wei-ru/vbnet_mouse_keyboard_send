
Imports System.Runtime.InteropServices
Imports System.IO
Imports System.Text

Public Class Form1
    Public Declare Ansi Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Integer
    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Integer, ByVal hWnd2 As Integer, ByVal lpsz1 As String, ByVal lpsz2 As String) As Integer

    Dim systemSend As New SystemSend()

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ListView1.View = System.Windows.Forms.View.Details
        ListView1.Columns.Add("Window Class or Title", 250)
        ListView1.Columns.Add("Handle")
        ListView1.Columns.Add("Type")
        ListView1.Columns.Add("Class Name", 150)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ListView1.Items.Clear()
        EnumWindows(AddressOf EnumWindowProc, &H0)


    End Sub

    Private Sub ListView1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListView1.SelectedIndexChanged

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
End Class

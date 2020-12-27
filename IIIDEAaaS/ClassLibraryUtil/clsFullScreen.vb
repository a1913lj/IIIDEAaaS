Imports System.Windows.Forms

Public Class classFullScreen
    Dim varScreen As Screen
    Dim intWidth As Integer = Screen.PrimaryScreen.Bounds.Width
    Dim intHeight As Integer = Screen.PrimaryScreen.Bounds.Height
    Dim intTop As Integer = 0
    Dim intLeft As Integer = 0
    Dim intX As Integer = 0
    Dim intY As Integer = 0

    Public Function FullscreenTheForm(ByVal frm As Form)
        frm.Top = intTop
        frm.Left = intLeft
        frm.Width = intWidth + 40
        frm.Height = intHeight
        frm.FormBorderStyle = FormBorderStyle.None
        Return 0
    End Function
End Class

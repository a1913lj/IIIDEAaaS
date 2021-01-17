Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim url As String = "https://chikaku-navi.com/harman/"
        url = Me.ComboBox1.Text

        Me.Enabled = False

        Try
            Me.TopMost = True
            Me.TopMost = False
            Dim p As Process = Process.Start(url)
            p.WaitForExit()
        Catch ex As Exception
        Finally
            Me.Enabled = True
            Me.WindowState = FormWindowState.Maximized
            Me.TopMost = True
            Me.TopMost = False
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Dispose()
    End Sub
End Class
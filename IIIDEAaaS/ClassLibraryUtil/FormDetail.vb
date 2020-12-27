Imports System.Windows.Forms

Public Class FormDetail
    Dim _bShow As Boolean = False

    Delegate Sub DelegateSetMessage(ByVal str As String)

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Me.Visible = Not Me.Visible
        Application.DoEvents()
    End Sub

    Public Sub SetMessage(Optional ByVal strMessage As String = "Alert")
        If Me.Label1.InvokeRequired Then
            Invoke(New DelegateSetMessage(AddressOf SetMessage), strMessage)
        Else
            Label1.Text = strMessage
        End If
    End Sub

    Public Sub tmr(Optional ByVal bStart As Boolean = True)

        If bStart Then
            Me.Timer1.Enabled = True
            Me.Timer1.Start()
            Me.Opacity = 0.5
            Me.Show()
        Else
            Me.Timer1.Stop()
            Me.Timer1.Enabled = False
            Me.Hide()
        End If
    End Sub
End Class
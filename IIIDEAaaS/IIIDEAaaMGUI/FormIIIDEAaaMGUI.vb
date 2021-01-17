Public Class FormIIIDEAaaMGUI

    Private m_Lock As New Object
    Private s As String = ""
    Dim m_thread1 As System.Threading.Thread = New System.Threading.Thread(New Threading.ParameterizedThreadStart(AddressOf Worker))
    Dim m_thread2 As System.Threading.Thread = New System.Threading.Thread(New Threading.ParameterizedThreadStart(AddressOf Worker))

    Private m_frm As ClassLibraryUtil.FormDetail = Nothing

    Enum Category
        idHerrmann = 1
        idTeam = 2
        idbrainstorming
        idGrouping
    End Enum


    Private Sub Worker(obj As Object)
        Dim lImage As Long = 1
        Dim dic1 As New Dictionary(Of Long, Bitmap) From {{1, IIIDEAaaMGUI.My.Resources.Resources._3},
            {2, IIIDEAaaMGUI.My.Resources.Resources._1},
            {3, IIIDEAaaMGUI.My.Resources.Resources._2},
            {4, IIIDEAaaMGUI.My.Resources.Resources._4}}

        Dim dic2 As New Dictionary(Of Long, Bitmap) From {{1, IIIDEAaaMGUI.My.Resources.Resources.i_kun01},
            {2, IIIDEAaaMGUI.My.Resources.Resources.i_kun01_Blue},
            {3, IIIDEAaaMGUI.My.Resources.Resources.i_kun_timer01}}

        Dim rnd As New Random()
        Do
            SyncLock m_Lock
                If obj Is PictureBox1 Then
                    CType(obj, PictureBox).Image = dic1(lImage)
                    If lImage >= dic1.Count Then
                        lImage = 1
                    Else
                        lImage += 1
                    End If
                Else
                    CType(obj, PictureBox).Image = dic2(lImage)
                    If lImage >= dic2.Count Then
                        lImage = 1
                    Else
                        lImage += 1
                    End If
                End If
            End SyncLock
            Randomize()
            Threading.Thread.Sleep(rnd.Next(1 * 1000, 3 * 1000))
        Loop While (True)
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        For Each k As KeyValuePair(Of Threading.Thread, PictureBox) In {
            New KeyValuePair(Of Threading.Thread, PictureBox)(m_thread1, PictureBox1),
            New KeyValuePair(Of Threading.Thread, PictureBox)(m_thread2, PictureBox2)}
            k.Key.Start(k.Value)
        Next

        SendToBack()
    End Sub

    Private Sub Enable(Optional ByVal bEnabled As Boolean = True)

        For Each ctrl In {Me.Button1, Me.Button2, Me.Button3, Me.Button4}
            ctrl.Enabled = bEnabled
        Next

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Enable(False)
        Me.SendToBack()
        Me.BackgroundWorker1.RunWorkerAsync(Category.idTeam)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Enable(False)
        Me.SendToBack()
        Me.BackgroundWorker1.RunWorkerAsync(Category.idGrouping)
    End Sub


    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork

        Dim emCategory As Category = e.Argument
        If emCategory = Category.idHerrmann Then
            Dim frm As New Form1
            frm.ShowDialog()
        ElseIf emCategory = Category.idTeam OrElse emCategory = Category.idGrouping Then
            Dim fn As String = IIf(emCategory = Category.idTeam, "IIIdeaaas2Teams.exe", "IIIDeaaas2Group.exe")
            Dim p As Process = Process.Start(fn)
            p.WaitForExit()
        ElseIf emCategory = Category.idbrainstorming Then
            DummyBrainstorming()
        End If
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Enable(True)
        If Not m_frm Is Nothing Then
            m_frm.Dispose()
            m_frm = Nothing
        End If

        Me.WindowState = FormWindowState.Normal
        Me.Activate()
        Me.TopMost = True
        Me.TopMost = False
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        For Each t As Threading.Thread In {m_thread1, m_thread2}
            t.Abort()
            t.Join()
        Next
    End Sub

    Private Sub ShowInfo(ByVal strInfo As String)
        'WEB
        If Not m_frm Is Nothing Then
            m_frm.Dispose()
            m_frm = Nothing
        End If

        m_frm = New ClassLibraryUtil.FormDetail
        m_frm.SetMessage(strInfo)
        m_frm.tmr(True)
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Enable(False)
        Me.SendToBack()
        ShowInfo("外部診断＞" & CType(sender, Button).Text & "...")
        Me.BackgroundWorker1.RunWorkerAsync(Category.idHerrmann)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Enable(False)
        Me.SendToBack()
        ShowInfo("オンライン＞" & CType(sender, Button).Text & "...")
        Me.BackgroundWorker1.RunWorkerAsync(Category.idbrainstorming)
    End Sub


    Private Sub DummyBrainstorming()
        Dim fn As String = "OnlineSample.pptx"
        Dim attributes As FileAttribute = IO.File.GetAttributes(fn)
        IO.File.SetAttributes(fn, attributes Or IO.FileAttributes.ReadOnly)

        Dim c As New ClassLibraryUtil.clsStartProcess(New ProcessStartInfo)
        With c.ProcessStartInfo
            .WindowStyle = ProcessWindowStyle.Maximized
            .Verb = "Open"
            .FileName = fn
        End With
        c.ExecSync()
        IO.File.SetAttributes(fn, attributes)
    End Sub
End Class

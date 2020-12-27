Public Class clsStartProcess
    Dim m_psInfo As New ProcessStartInfo
    Dim m_Process As Process = Nothing

    Public Sub New(psInfo As ProcessStartInfo)
        m_psInfo = psInfo
    End Sub

    Public Sub New(ByVal FileName As String,
                   Optional ByVal Arguments As String = "",
                   Optional CreateNoWindow As Boolean = False,
                   Optional RedirectStandardOutput As Boolean = True,
                   Optional UseShellExecute As Boolean = False)
        With m_psInfo
            .FileName = FileName '    // 実行するファイル 
            .WorkingDirectory = System.Environment.CurrentDirectory
            If String.IsNullOrEmpty(Arguments) = False Then
                .Arguments = Arguments '    // コマンドパラメータ（引数）
            End If
            .CreateNoWindow = CreateNoWindow '   // コンソール・ウィンドウを開かない
            .UseShellExecute = UseShellExecute '  // シェル機能を使用しない
            .RedirectStandardOutput = RedirectStandardOutput
        End With
    End Sub

    Public Property ProcessStartInfo() As ProcessStartInfo
        Get
            Return m_psInfo
        End Get
        Set(ByVal value As ProcessStartInfo)
            m_psInfo = value
        End Set
    End Property


    Public Function ExecSync() As String
        m_Process = Process.Start(m_psInfo)
        m_Process.WaitForExit()
        If m_psInfo.RedirectStandardOutput Then
            Return m_Process.StandardOutput.ReadToEnd().Replace(vbCr + vbCrLf, vbLf)
        End If
        Return ""
    End Function
End Class

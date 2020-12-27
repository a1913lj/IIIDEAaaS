Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Diagnostics
''' <summary>
''' Python呼び出す制御
''' </summary>
''' 
Public Class clsCallPython
    Dim m_pythonPath As String = "%ProgramData%\Anaconda3\Python.exe"
    ''' <summary>
    ''' コンストラクター
    ''' </summary>
    Public Sub New()
        m_pythonPath = Environment.ExpandEnvironmentVariables(m_pythonPath)

        If IO.File.Exists(m_pythonPath) = False Then
            m_pythonPath = Environment.ExpandEnvironmentVariables("%USERPROFILE%\Anaconda3\Python.exe")
        End If
    End Sub
    ''' <summary>
    ''' プロセス実行ルーチン、実行結果のリダイレクト情報を返す
    ''' </summary>
    ''' <param name="pythonFilePath">Pythonスクリプトファイル</param>
    ''' <param name="param">スクリプト引数</param>
    ''' <returns></returns>
    ''' 
    Public Function callExecute(ByVal pythonFilePath As String, Optional ByVal param As List(Of String) = Nothing) As String
        Dim lRet As Long = 0L
        Dim paramList = New List(Of String)
        paramList.Add("""" & pythonFilePath & """")
        If Not param Is Nothing Then
            For Each s As String In param
                paramList.Add("""" & s & """")
            Next
        End If

        'プロセス実行情報詳細設定
        Dim psi = New ProcessStartInfo(m_pythonPath)
        With psi
            .FileName = m_pythonPath
            .Arguments = Strings.Join(paramList.ToArray, " ")
            .CreateNoWindow = True
            .WindowStyle = ProcessWindowStyle.Hidden
            .WorkingDirectory = System.Environment.CurrentDirectory
            .RedirectStandardOutput = True
            .UseShellExecute = False
        End With
        Dim ps = Process.Start(psi)
        ps.WaitForExit()
        lRet = ps.ExitCode
        'リダイレクト結果
        Dim output As String = ps.StandardOutput.ReadToEnd().Replace(vbCr + vbCrLf, vbLf)
        ps.Close()
        GC.Collect()
        Return output
    End Function
End Class

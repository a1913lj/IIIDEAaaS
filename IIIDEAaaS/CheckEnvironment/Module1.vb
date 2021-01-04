Imports System.IO
Imports System.Windows.Forms
Imports System.Text.RegularExpressions

Module Module1

    Sub Main()
        Do
            Console.WriteLine(System.Reflection.MethodBase.GetCurrentMethod.Name)
            If CheckAnaconda() <> 0 Then
                Exit Do
            End If

            For Each s As String In {"C:\ProgramData\Anaconda3",
                "C:\ProgramData\Anaconda3\Library\mingw-w64\bin",
                "C:\ProgramData\Anaconda3\Library\usr\bin",
                "C:\ProgramData\Anaconda3\Library\bin"}
                AddSystemEnvironmentPath(s)
            Next

            If CheckMeCab() <> 0 Then
                Exit Do
            End If

            AddSystemEnvironmentPath("C:\Program Files\MeCab\bin")

            If CheckNLP() <> 0 Then
                Exit Do
            End If

        Loop While (False)

    End Sub


    Private Function CheckAnaconda(Optional ByVal bChkOnly As Boolean = True) As Long
        Console.WriteLine(System.Reflection.MethodBase.GetCurrentMethod.Name)
        Dim lRet As Long = 0
        Dim bat1 As String = "%ProgramData%\Anaconda3\Scripts\activate.bat"
        Dim bat2 As String = "%USERPROFILE%\Anaconda3\Scripts\activate.bat"

        Dim bIns As Boolean = False
        For Each bat As String In {bat1, bat2}
            bat = System.Environment.ExpandEnvironmentVariables(bat)
            If IO.File.Exists(bat) Then
                bIns = True
                Exit For
            End If
        Next

        If bIns = False Then
            If MessageBox.Show("Anacondaはインストールされていません。" & vbCrLf & "Anacondaのインストールをしますか", "", MessageBoxButtons.YesNo, MessageBoxIcon.Information) = DialogResult.Yes Then
                Dim fn As String = ".\Anaconda3\Anaconda3-2020.11-Windows-x86_64.exe"
                StartProc(fn)
            End If
        End If

        Return IIf(bIns, 0, -99)
    End Function

    Private Function CheckNLP() As Long

        'pip install janome
        'pip install sumy
        'pip install tinysegmenter
        'pip install pysummarization
        'pip install MeCab

        Console.WriteLine(System.Reflection.MethodBase.GetCurrentMethod.Name)

        Dim lst As New List(Of String) From {"janome", "sumy", "tinysegmenter", "pysummarization", "MeCab"}

        For Each s As String In lst
            Dim lRet As Long = StartProc("pip", "show " & """" & s & """", True)
            If lRet <> 0 Then
                StartProc("pip", "install" & " " & """" & s & """", False)
            End If
        Next
        Return 0
    End Function

    Private Function CopyMeCab() As Long
        Console.WriteLine(System.Reflection.MethodBase.GetCurrentMethod.Name)
        Dim lRet As Long = -1
        Dim src As String = "C:\Program Files\MeCab\bin\libmecab.dll"
        If IO.File.Exists(src) Then
            For Each dir As String In {"%ProgramData%\Anaconda3", "%USERPROFILE%\Anaconda3"}
                dir = System.Environment.ExpandEnvironmentVariables(dir)
                If IO.Directory.Exists(dir) Then
                    Dim dll As String = "Lib\site-packages\libmecab.dll"
                    Dim dst As String = IO.Path.Combine(dir, dll)
                    If IO.File.Exists(dst) = False Then
                        Dim fa As IO.FileAttributes = IO.File.GetAttributes(dst)
                        fa = fa And Not FileAttributes.ReadOnly
                        File.SetAttributes(dst, fa)
                        File.Copy(src, dst, True)
                    End If
                    lRet = 0
                End If
            Next
        End If
        Return lRet
    End Function
    Private Function CheckMeCab() As Long
        Console.WriteLine(System.Reflection.MethodBase.GetCurrentMethod.Name)
        Dim src As String = "C:\Program Files\MeCab\bin\libmecab.dll"

        If IO.File.Exists(src) = False Then
            Dim fn As String = ".\NLP\mecab-0.996-64.exe"
            StartProc(fn)
        End If
        Return CopyMeCab()
    End Function

    Private Function StartProc(ByVal fn As String, ByVal Arguments As List(Of String), Optional ByVal bDQM As Boolean = False, Optional ByVal bCreateNoWindow As Boolean = False)
        Console.WriteLine(System.Reflection.MethodBase.GetCurrentMethod.Name)
        Dim args As String = ""

        If bDQM Then
            For i As Int32 = 0 To Arguments.Count - 1
                Arguments(i) = """" & Arguments(i) & """"
            Next
        End If
        args = Join(Arguments.ToArray, " ")
        Return StartProc(fn, args, bCreateNoWindow)
    End Function

    Private Function StartProc(ByVal fn As String, Optional ByVal Arguments As String = "", Optional ByVal bCreateNoWindow As Boolean = False)
        Console.WriteLine("{0}->{1} {2}", System.Reflection.MethodBase.GetCurrentMethod.Name, fn, Arguments)
        Dim psi As New System.Diagnostics.ProcessStartInfo()
        With psi
            .FileName = fn
            .Arguments = Arguments
            .CreateNoWindow = bCreateNoWindow
            .UseShellExecute = True
            If bCreateNoWindow Then
                .WindowStyle = ProcessWindowStyle.Hidden
            End If
        End With
        Dim p As System.Diagnostics.Process = System.Diagnostics.Process.Start(psi)
        p.Start()
        p.WaitForExit()
        Return p.ExitCode
    End Function

    Private Function IncludeSystemEnvironmentPath(ByVal strPath As String) As Boolean
        Console.WriteLine("{0}->{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, strPath)
        Dim bInc As Boolean = False
        strPath = Regex.Replace(strPath, "[\/]+", "\")
        For Each s As String In Environment.GetEnvironmentVariable("Path", EnvironmentVariableTarget.Machine).Split(";")
            s = Regex.Replace(System.Environment.ExpandEnvironmentVariables(s), "[\/]+", "\").Trim
            If String.IsNullOrEmpty(s) Then Continue For
            If IO.Directory.Exists(s) = False Then Continue For
            If String.Compare(strPath, s, True) = 0 Then
                bInc = True
                Exit For
            End If
        Next
        Return bInc
    End Function
    Private Sub AddSystemEnvironmentPath(ByVal strPath As String)
        Console.WriteLine("{0}->{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, strPath)
        strPath = strPath.Trim
        If String.IsNullOrEmpty(strPath) = False AndAlso IO.Directory.Exists(strPath) = True AndAlso IncludeSystemEnvironmentPath(strPath) = False Then
            Console.WriteLine("{0}->{1}", System.Reflection.MethodBase.GetCurrentMethod.Name, strPath)
            System.Environment.SetEnvironmentVariable("Path", System.Environment.GetEnvironmentVariable("Path", EnvironmentVariableTarget.Machine) & ";" & strPath)
        End If
    End Sub
End Module

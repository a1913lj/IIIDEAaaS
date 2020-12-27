Imports System.Drawing
Imports System.Windows.Forms
''' <summary>
''' ユーティリティー
''' </summary>
''' 
Public Class clsUtil
    ''' <summary>
    ''' ファイル読み込み結果を返す
    ''' </summary>
    ''' <param name="fn">ファイルパス</param>
    ''' <param name="en">エンコーディング</param>
    ''' <returns>ファイル読み込み結果を返す</returns>
    ''' 
    Public Shared Function ReadFile(ByVal fn As String, en As System.Text.Encoding) As List(Of String)
        Dim lst As New List(Of String)
        Do
            If IO.File.Exists(fn) = False Then Exit Do
            Using sr As New System.IO.StreamReader(fn, en)
                Do While sr.Peek() >= 0
                    Dim strLine As String = sr.ReadLine().Trim
                    If String.IsNullOrEmpty(strLine) Then Continue Do
                    lst.Add(strLine)
                Loop
                sr.Close()
            End Using
        Loop While (False)
        Return lst
    End Function
    ''' <summary>
    ''' ファイル読み込み結果を返す
    ''' </summary>
    ''' <param name="fn">ファイル名</param>
    ''' <returns>ファイル読み込み結果を返す</returns>
    ''' 
    Public Shared Function ReadFile(ByVal fn As String) As List(Of String)
        Dim lst As New List(Of String)
        Do
            If IO.File.Exists(fn) = False Then Exit Do
            'Using sr As New System.IO.StreamReader(fn, System.Text.Encoding.UTF8)
            'System.Text.Encoding.GetEncoding("Shift-JIS")
            Using sr As New System.IO.StreamReader(fn)
                Do While sr.Peek() >= 0
                    Dim strLine As String = sr.ReadLine().Trim
                    If String.IsNullOrEmpty(strLine) Then Continue Do
                    lst.Add(strLine)
                Loop
                sr.Close()
            End Using
        Loop While (False)
        Return lst
    End Function

    ''Shift JISとして文字列に変換
    'str = System.Text.Encoding.GetEncoding(932).GetString(bytesData)
    ''JISとして変換
    'str = System.Text.Encoding.GetEncoding(50220).GetString(bytesData)
    ''EUCとして変換
    'str = System.Text.Encoding.GetEncoding(51932).GetString(bytesData)
    ''UTF-8として変換
    'str = System.Text.Encoding.UTF8.GetString(bytesData)
    ''' <summary>
    ''' ファイル書き込み
    ''' </summary>
    ''' <param name="strData">書き込みデータ</param>
    ''' <param name="fn">書き込みファイル名</param>
    ''' <param name="bAppend">追加/新規：true/false</param>
    ''' <returns>0/0以外：OK/NG</returns>
    Public Shared Function WriteFile(ByVal strData As String, ByVal fn As String, Optional ByVal bAppend As Boolean = False) As Long
        Dim lRet As Long = 0
        Do
            'Using sw As New System.IO.StreamWriter(fn, bAppend, System.Text.Encoding.UTF8)
            Using sw As New System.IO.StreamWriter(fn, bAppend)
                sw.Write(strData)
                sw.Close()
            End Using
        Loop While (False)
        Return lRet
    End Function
    ''' <summary>
    ''' ファイル書き込み
    ''' </summary>
    ''' <param name="lst">書き込みデータ</param>
    ''' <param name="fn">書き込みファイル名</param>
    ''' <param name="bAppend">追加/新規：true/false</param>
    ''' <returns>0/0以外：OK/NG</returns>
    ''' 
    Public Shared Function WriteFile(ByVal lst As List(Of String), ByVal fn As String, Optional ByVal bAppend As Boolean = False) As Long
        Return WriteFile(Join(lst.ToArray, vbCrLf), fn, bAppend)
    End Function
    ''' <summary>
    ''' ファイル書き込み
    ''' </summary>
    ''' <param name="sb">書き込みデータ</param>
    ''' <param name="fn">書き込みファイル名</param>
    ''' <param name="bAppend">追加/新規：true/false</param>
    ''' <returns>0/0以外：OK/NG</returns>
    '''
    Public Shared Function WriteFile(ByVal sb As System.Text.StringBuilder, ByVal fn As String, Optional ByVal bAppend As Boolean = False) As Long
        Return WriteFile(sb.ToString, fn, bAppend)
    End Function

    ''' <summary>
    ''' ディレクトリ作成
    ''' </summary>
    ''' <param name="strPath">ディレクトリ名</param>
    ''' <returns>0/0以外：OK/NG</returns>
    ''' 
    Public Shared Function CreateDirectory(ByVal strPath As String) As Long
        Dim lRet As Long = 0
        If IO.Directory.Exists(strPath) = False Then
            IO.Directory.CreateDirectory(strPath)
        End If
        Return lRet
    End Function
    ''' <summary>
    ''' イメージ情報作成
    ''' </summary>
    ''' <param name="filename">画像ファイル名</param>
    ''' <returns>画像ファイルのイメージ情報</returns>
    ''' 
    Public Shared Function CreateImage(ByVal filename As String) As System.Drawing.Image
        Dim fs As New System.IO.FileStream(filename, System.IO.FileMode.Open, System.IO.FileAccess.Read)
        Dim img As System.Drawing.Image = System.Drawing.Image.FromStream(fs)
        fs.Close()
        Return img
    End Function

    ''' <summary>
    ''' カレントプロセス名取得
    ''' </summary>
    ''' <returns>プロセスイメージ名</returns>
    ''' 
    Public Shared Function ProcessName() As String
        Return IO.Path.GetFileNameWithoutExtension(System.Diagnostics.Process.GetCurrentProcess().ProcessName)
    End Function

    ''' <summary>
    ''' ファイル開くダイアログ
    ''' </summary>
    ''' <param name="filter">フィルター</param>
    ''' <param name="title">タイトル</param>
    ''' <returns>開く対象ファイル名</returns>
    ''' 
    Public Shared Function OpenFileDialog(Optional ByVal filter As String = "", Optional ByVal title As String = "") As String
        Dim strFile As String = ""
        Const _FILTER As String = "すべてのファイル(*.*)|*.*"
        Const _TITLE As String = "ファイルを選択してください"
        Do
            Dim ofd As New OpenFileDialog
            With ofd
                .Filter = IIf(String.IsNullOrEmpty(filter), _FILTER, filter)
                .Title = IIf(String.IsNullOrEmpty(title), _TITLE, title)
                .FilterIndex = 1
                .CheckFileExists = True
            End With
            If ofd.ShowDialog <> Windows.Forms.DialogResult.OK Then Exit Do
            strFile = ofd.FileName
        Loop While (False)

        Return strFile
    End Function



End Class

''' <summary>
''' WIN32メソッドオーバーラップ
''' </summary>
''' 
Public Class Win32

    <Runtime.InteropServices.DllImport("kernel32.dll")> Public Shared Function AllocConsole() As Boolean

    End Function
    <Runtime.InteropServices.DllImport("kernel32.dll")> Public Shared Function FreeConsole() As Boolean

    End Function


    <Runtime.InteropServices.DllImport("USER32.dll")> Public Shared Function FindWindow(ByVal lpClassName As String, ByVal lpWindowName As String) As System.IntPtr

    End Function

    <Runtime.InteropServices.DllImport("USER32.dll")> Public Shared Function SetParent(ByVal hWndChild As System.IntPtr, ByVal hWndNewParent As System.IntPtr) As System.IntPtr

    End Function

    <Runtime.InteropServices.DllImport("USER32.dll")> Public Shared Function GetDC(ByVal hwnd As System.IntPtr) As System.IntPtr

    End Function

    <Runtime.InteropServices.DllImport("USER32.dll")> Public Shared Function ReleaseDC(ByVal hwnd As System.IntPtr, ByVal hdc As System.IntPtr) As System.IntPtr

    End Function
End Class

'Public Class clsDesktop
'    Implements IDisposable

'    Protected _disposed = False
'    Private m_thisLock As New Object
'    Private m_strMeaasge As String = ""
'    Private m_lViewTime As Long = 500

'    Dim m_thread As System.Threading.Thread = Nothing
'    Dim m_bStop As Boolean = False
'    Public Overloads Sub Dispose() Implements IDisposable.Dispose
'        Dispose(True)
'    End Sub

'    Protected Overrides Sub Finalize()
'        MyBase.Finalize()
'        Dispose(False)
'    End Sub

'    Protected Overridable Overloads Sub Dispose(ByVal disposing As Boolean)
'        If Not _disposed Then
'            If disposing Then
'                '*** アンマネージリソースの開放
'            End If
'            '*** マネージドリソースの開放
'            Me.Sync()

'        End If
'        _disposed = True
'    End Sub

'    Public Sub New(ByVal strMeaasge As String)
'        m_strMeaasge = strMeaasge
'        m_thread = New System.Threading.Thread(AddressOf DrawStringDesktop)
'        m_thread.Start()
'    End Sub

'    Public Sub Sync()
'        m_bStop = True
'        m_thread.Join()
'    End Sub

'    Private Sub DrawStringDesktop()

'        Dim b As Boolean = True
'        Dim disDC As System.IntPtr = Win32.GetDC(IntPtr.Zero)
'        Dim bmp As System.Drawing.Bitmap = New System.Drawing.Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height)
'        Dim font As Font = New Font("MSゴシック", 48, FontStyle.Bold)

'        Dim g As Graphics = Nothing
'        g = Graphics.FromHdc(disDC)
'        Dim hDC As System.IntPtr = g.GetHdc
'        g.ReleaseHdc(hDC)
'        Dim stringSize As SizeF = g.MeasureString(m_strMeaasge, font, 1000)
'        g.DrawString(m_strMeaasge, font, New SolidBrush(Color.Red), (Screen.PrimaryScreen.Bounds.Width / 2) - (stringSize.Width / 2), 0)
'        If g IsNot Nothing Then g.Dispose()
'        GC.Collect()
'        Win32.ReleaseDC(IntPtr.Zero, disDC)
'    End Sub
'End Class

Public Class clsLabeling
    ''' <summary>
    ''' ラベリングクラスメソッド
    ''' </summary>
    ''' 
    Dim m_py As String = "janome_ver0.10.0.py"
    Dim m_fn As String = ""
    Dim m_strCategory As String = "MeCab"

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="fn">グルーピングの処理結果</param>
    ''' <param name="py">ラベリングのＰｙｔｈｏｎコード</param>
    ''' <param name="strCategory">ＮＬＰラベリング：MeCab/Janome</param>
    ''' 
    Public Sub New(ByVal fn As String, Optional ByVal py As String = "janome_ver0.10.0.py", Optional ByVal strCategory As String = "MeCab")
        If IO.File.Exists(fn) Then m_fn = fn
        If IO.File.Exists(py) Then m_py = py
        m_strCategory = strCategory
    End Sub


    Public Property Category() As String
        Get
            Return m_strCategory
        End Get
        Set(ByVal value As String)
            m_strCategory = value
        End Set
    End Property

    ''' <summary>
    ''' ラベリングクラスメソッド
    ''' NLP処理リダイレクト情報を返す
    ''' </summary>
    ''' <param name="rst"></param>
    ''' <returns>NLPの処理結果を返す</returns>
    ''' 
    Public Function Labeling(ByVal rst As String) As String
        Dim py As New clsCallPython
        Dim hProcess As System.Diagnostics.Process = System.Diagnostics.Process.GetCurrentProcess()
        Dim strEnv As String = hProcess.SessionId.ToString("X4") & hProcess.Id.ToString("X4")

        Environment.SetEnvironmentVariable(strEnv, "XXX", EnvironmentVariableTarget.Process)
        Return py.callExecute(m_py, New List(Of String) From {m_fn, rst})
    End Function
End Class

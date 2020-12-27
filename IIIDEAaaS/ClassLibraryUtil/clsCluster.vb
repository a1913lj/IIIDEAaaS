''' <summary>
''' グルーピング
''' </summary>
Public Class clsCluster
    Implements IDisposable

    ''' <summary>
    ''' MS POWERPOINT SHPE情報
    ''' </summary>
    ''' 
    Public Class clsShape
        Private _idx As Long = 0
        Private _ShapeName As String = 0
        Private _ShapeText As String = ""

        Private _Width As Single = 0.0#
        Private _Height As Single = 0.0#


        Public Sub New(ByVal idx As Long, ByVal ShapeName As String, Optional shapeText As String = "",
                       Optional Width As Single = 0.0#, Optional Height As Single = 0.0#)
            _idx = idx
            _ShapeName = ShapeName
            _ShapeText = shapeText
            _Width = Width
            _Height = Height
        End Sub

        Public Property index() As Long
            Get
                Return _idx
            End Get
            Set(ByVal value As Long)
                _idx = value
            End Set
        End Property

        Public Property Shape() As String
            Get
                Return _ShapeName
            End Get
            Set(ByVal value As String)
                _ShapeName = value
            End Set
        End Property

        Public Property Text() As String
            Get
                Return _ShapeText
            End Get
            Set(ByVal value As String)
                _ShapeText = value
            End Set
        End Property

        Public Property Width() As Single
            Get
                Return _Width
            End Get
            Set(ByVal value As Single)
                _Width = value
            End Set
        End Property

        Public Property Height() As Single
            Get
                Return _Height
            End Get
            Set(ByVal value As Single)
                _Height = value
            End Set
        End Property
    End Class

    Public Overloads Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
        Dispose(False)
    End Sub

    Protected Overridable Overloads Sub Dispose(ByVal disposing As Boolean)
        If Not _disposed Then
            If disposing Then
                '*** アンマネージリソースの開放

            End If
            '*** マネージドリソースの開放

        End If
        _disposed = True
    End Sub

    Protected _disposed = False


    Dim m_py As String = "cluster_ver.0.10.0.py"
    Dim m_fn As String = ""
    Dim m_cluster As Long = 5
    Public Sub New(ByVal fn As String, Optional ByVal cluster As Long = 5, Optional ByVal py As String = "cluster_ver.0.10.0.py")
        If IO.File.Exists(fn) Then m_fn = fn
        m_cluster = cluster
        If IO.File.Exists(py) Then m_py = py
    End Sub


    'SlideNo.,ClusterNo.,IndexNo.,ShapeObjectType,Text
    '1,0,19,Rectangle 131,さぁ アイデアだせ
    '1,0,22,Rectangle 44,アイデア具現化無双
    '1,0,27,Rectangle 47,初心者でも使える簡単アイデアソリューション
    '1,1,6,Rectangle 37,逆アイデア から攻め
    '1,1,8,Rectangle 40,iTPS (Idea Total solution package)
    '1,1,9,Rectangle 41,Idea Institute
    '1,1,12,Rectangle 52,SMART IDEA

    ''' <summary>
    ''' クラスタリングPyhtoのNLP処理処理結果を一覧を返す
    ''' </summary>
    ''' <param name="fn">NLP処理結果ファイルパス</param>
    ''' <returns>クラスタ一覧を返す</returns>
    ''' 
    Private Function Clutser(ByVal fn As String) As Dictionary(Of Int32, Dictionary(Of Int32, List(Of clsShape)))
        Dim dic As New Dictionary(Of Int32, Dictionary(Of Int32, List(Of clsShape)))

        For Each s As String In clsUtil.ReadFile(fn)
            s = s.Trim
            If String.IsNullOrEmpty(s) Then Continue For
            Dim ss() As String = s.Split(",")
            If ss.Length < 4 Then Continue For
            'SlideNo.,ClusterNo.,IndexNo.,ShapeObjectType,Text
            Dim slide As Long = 0
            Dim cluster As Long = 0
            Dim idx As Long = 0
            Dim shape As String = ss(3)
            If Long.TryParse(ss(0), slide) = False OrElse slide <= 0 Then Continue For
            If Long.TryParse(ss(1), cluster) = False OrElse cluster < 0 OrElse cluster > m_cluster Then Continue For
            If Long.TryParse(ss(2), idx) = False OrElse idx < 0 Then Continue For
            Dim lst As New List(Of String)
            For i As Int32 = 4 To ss.Length - 1
                lst.Add(ss(i))
            Next

            If dic.ContainsKey(slide) = False Then
                dic.Add(slide, New Dictionary(Of Integer, List(Of clsShape)))
            End If

            If dic(slide).ContainsKey(cluster) = False Then
                dic(slide).Add(cluster, New List(Of clsShape))
            End If
            dic(slide)(cluster).Add(New clsShape(idx, shape, Join(lst.ToArray, ",")))
        Next
        Return dic
    End Function
    ''' <summary>
    ''' クラスタリングPyhtoのNLP処理処理結果を一覧を返す
    ''' </summary>
    ''' <param name="bXmlFile">NLP処理結果ファイルパス</param>
    ''' <returns>クラスタ一覧を返す</returns>
    ''' 
    Public Function Clutser(Optional ByVal bXmlFile As Boolean = True) As Dictionary(Of Int32, Dictionary(Of Int32, List(Of clsShape)))
        Dim py As New clsCallPython
        py.callExecute(m_py, New List(Of String) From {m_fn, m_cluster.ToString, IIf(bXmlFile, "xml", "txt")})
        Dim rst As String = IO.Path.Combine(IO.Path.GetDirectoryName(m_fn), String.Concat(IO.Path.GetFileNameWithoutExtension(m_fn), ".rst"))
        Return Clutser(rst)
    End Function
End Class

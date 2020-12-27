
''' <summary>
''' 技術的これはA象限に属する描写語です。A象限では、論理的でシステマチックな知的処理が行われ、事実や数字、統計など目に見えるものを重視します。また、データの裏付けがあり、先例のある結論を好みます。
'''緻密なこれはB象限に属する描写語です。B象限では、実践や手続きを重視し、能率や秩序、規律への傾向を表し、課題をシステマチックに順序立てて処理完遂します。時間は有効的に管理されます。
'''友好的これはC象限に属する描写語です。C象限では、ムード、雰囲気、態度を重視し、感受性があり、受容的です。人間に関心が高く、自己表現が上手です。
'''概念的これはD象限に属する描写語です。D象限では、メンタルなインプットを同時にすばやく関連付けることができ、抽象的な概念を心地よく感じます。また、問'''題解決には初めから全体的なアプローチをとります。
''' 
''' </summary>

Public Class clsHerrmann
    Dim _name As String = ""
    Dim _a As Long = 0
    Dim _b As Long = 0
    Dim _c As Long = 0
    Dim _d As Long = 0

    Public Sub New(ByVal name As String, Optional ByVal a As Long = 50L, Optional ByVal b As Long = 50L, Optional ByVal c As Long = 50L, Optional ByVal d As Long = 50L)
        _name = name
        _a = a
        _c = b
        _c = c
        _d = d
    End Sub

    Public Property N() As String
        Get
            Return _name
        End Get
        Set(ByVal value As String)
            _name = value
        End Set
    End Property

    Public Property A() As Long
        Get
            Return _a
        End Get
        Set(ByVal value As Long)
            _a = value
        End Set
    End Property

    Public ReadOnly Property AN() As String
        Get
            Return "A【技術的】"
        End Get
    End Property
    Public ReadOnly Property BN() As String
        Get
            Return "B【緻密な】"
        End Get
    End Property
    Public ReadOnly Property CN() As String
        Get
            Return "C【友好的】"
        End Get
    End Property
    Public ReadOnly Property DN() As String
        Get
            Return "D【概念的】"
        End Get
    End Property

    Public Property B() As Long
        Get
            Return _b
        End Get
        Set(ByVal value As Long)
            _b = value
        End Set
    End Property

    Public Property C() As Long
        Get
            Return _c
        End Get
        Set(ByVal value As Long)
            _c = value
        End Set
    End Property

    Public Property D() As Long
        Get
            Return _d
        End Get
        Set(ByVal value As Long)
            _d = value
        End Set
    End Property

End Class


Class clsTeam
    Dim m_TeamName As String = ""
    Dim m_Menbers As New List(Of clsHerrmann)

    Public Sub New(ByVal TeamName As String)
        m_TeamName = TeamName
    End Sub

    Public Property TeamName() As String
        Get
            Return m_TeamName
        End Get
        Set(ByVal value As String)
            m_TeamName = value
        End Set
    End Property

    Public Property Menbers() As List(Of clsHerrmann)
        Get
            Return m_Menbers
        End Get
        Set(ByVal value As List(Of clsHerrmann))
            m_Menbers = value
        End Set
    End Property

    Public ReadOnly Property Facilitation() As String
        Get
            Dim strFacilitation As String = ""
            Dim l As Long = 0
            For Each h In m_Menbers
                If h.D >= l Then
                    strFacilitation = h.N
                    l = h.D
                End If
            Next
            Return strFacilitation
        End Get
    End Property
End Class

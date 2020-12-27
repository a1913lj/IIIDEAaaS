Imports System.IO
Imports System.Xml
Imports Microsoft.Office.Interop

''' <summary>
''' グルーピングメイン画面
''' </summary>
Public Class FormIIIDEAaaS2Group

    Const m_Opacity As Single = 0.7
    Private m_oPowerPoint As New clsPowerPoint
    Private m_dicCluster As New Dictionary(Of Int32, Dictionary(Of Int32, List(Of ClassLibraryUtil.clsCluster.clsShape))) 'グルーピング
    Private m_dicLabeling As New Dictionary(Of Int32, Dictionary(Of Int32, KeyValuePair(Of String, List(Of ClassLibraryUtil.clsCluster.clsShape)))) 'ラベリング
    Private m_frm As ClassLibraryUtil.FormDetail = Nothing
    Delegate Sub delegate_Detail(ByVal txt As TextBox, ByVal str As String)
    ''' <summary>
    ''' PPTX選択
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' 
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Detail(Join({System.Reflection.MethodBase.GetCurrentMethod.Name, " File Selecting..."}, ""))
        Const _FILTER As String = "ワークショップの結果ファイル(*.pptx)|*.pptx|すべてのファイル(*.*)|*.*"
        Const _TITLE As String = "アイデア出ししたPPTXのファイルを選択してください"
        Me.TextBox1.Text = ClassLibraryUtil.clsUtil.OpenFileDialog(_FILTER, _TITLE)
        Dim fn As String = TextBox1.Text
        If IO.File.Exists(fn) Then
            Detail(Join({fn}, "..."))
            If m_oPowerPoint.RdOpen(fn) Then
                'pptxを最前面
                Me.TopMost = True
                Me.TopMost = False
                Me.Activate()
                Enable(True)
            Else
                Enable(False)
            End If
        End If
    End Sub

    ''' <summary>
    ''' ボタンの活性化
    ''' </summary>
    ''' <param name="bEnable"></param>
    ''' <param name="objs"></param>
    ''' 
    Private Sub Enable(Optional ByVal bEnable As Boolean = True, Optional ByRef objs() As Control = Nothing)
        If objs Is Nothing Then
            objs = {Button1, ButtonLabeling, ButtonGrouping, ComboBox1, ComboBox2, ComboBox3}
        End If
        For Each obj In objs
            obj.Enabled = bEnable
        Next
    End Sub

    ''' <summary>
    ''' ボタンの活性化
    ''' </summary>
    ''' <param name="fn"></param>
    ''' <returns>PPTXの入力情報</returns>
    Private Function RdPptx(fn As String) As Dictionary(Of Int32, List(Of KeyValuePair(Of String, String)))
        Detail(Join({System.Reflection.MethodBase.GetCurrentMethod.Name, fn}, " > "))
        Return m_oPowerPoint.RdPptx(fn)
    End Function

    ''' <summary>
    ''' ログ出力
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="strText"></param>
    ''' 
    Private Sub Detail(sender As Object, strText As String)
        Dim TextBox As TextBox = DirectCast(sender, TextBox)
        With TextBox
            .Text += Now.ToString("yyyy/MM/dd HH:mm:ss.fff ") & strText & vbCrLf
            .SelectionStart = .TextLength
            .ScrollToCaret()
        End With

    End Sub
    ''' <summary>
    ''' ログ出力
    ''' </summary>
    ''' <param name="strText"></param>
    ''' 
    Private Sub Detail(strText As String)
        If TextBox2.InvokeRequired Then
            Invoke(New delegate_Detail(AddressOf Detail), TextBox2, strText)
        Else
            Detail(TextBox2, strText)
        End If

    End Sub

    ''' <summary>
    ''' XMLエレメント作成
    ''' </summary>
    ''' <param name="xmlDocument"></param>
    ''' <param name="key"></param>
    ''' <returns>XmlElement</returns>
    Private Function CreateElement(xmlDocument As XmlDocument, key As KeyValuePair(Of String, String)) As XmlElement
        Return CreateElement(xmlDocument, key.Key, key.Value)
    End Function

    ''' <summary>
    ''' XMLエレメント作成
    ''' </summary>
    ''' <param name="xmlDocument"></param>
    ''' <param name="strName"></param>
    ''' <param name="strValue"></param>
    ''' <returns>XmlElement</returns>
    Private Function CreateElement(ByVal xmlDocument As XmlDocument, ByVal strName As String, Optional ByVal strValue As String = "") As XmlElement
        Detail(Join({strName, strValue}, "="))
        Dim ele As XmlElement = xmlDocument.CreateElement(strName)
        If String.IsNullOrEmpty(strValue) = False Then
            Dim node As XmlNode = xmlDocument.CreateNode(XmlNodeType.Text, "", "")
            node.Value = strValue
            ele.AppendChild(node)
        End If
        Return ele
    End Function

    ''' <summary>
    ''' XMLエレメント作成
    ''' </summary>
    ''' <param name="strExtension"></param>
    ''' <returns>ファイル名を返す</returns>
    Private Function GetClusterFileName(Optional ByVal strExtension As String = ".xml") As String
        Dim fn As String = Me.TextBox1.Text
        'Dim strFile As String = String.Concat(IO.Path.GetFileNameWithoutExtension(IO.Path.GetRandomFileName()), ".xml")
        Dim strFile As String = IO.Path.Combine(IO.Path.GetDirectoryName(fn), String.Concat(IO.Path.GetFileNameWithoutExtension(fn), strExtension))
        strFile = IO.Path.Combine(Environment.CurrentDirectory, strFile)
        Return strFile
    End Function

    ''' <summary>
    ''' >PPTXのテキストボックスの情報を文字列生成、クラスリング用
    ''' </summary>
    ''' <param name="dic"></param>
    ''' <returns>PPTXのテキストボックスの情報を文字列を返す</returns>
    Private Function Rectangle2Text(ByVal dic As Dictionary(Of Int32, List(Of KeyValuePair(Of String, String)))) As String
        Dim strFile As String = GetClusterFileName(".txt")
        Dim lst As New List(Of String)
        For Each key1 As KeyValuePair(Of Int32, List(Of KeyValuePair(Of String, String))) In dic

            If key1.Value.Count = 0 Then Continue For
            For Each key2 As KeyValuePair(Of String, String) In key1.Value
                lst.Add(Join({key1.Key.ToString, key2.Key, key2.Value}, ","))
            Next
        Next
        ClassLibraryUtil.clsUtil.WriteFile(lst, strFile)
        Return strFile
    End Function

    ''' <summary>
    ''' >PPTXのテキストボックスの情報を文字列生成、クラスリング用
    ''' </summary>
    ''' <param name="dic"></param>
    ''' <returns>PPTXのテキストボックスの情報を文字列を返す</returns>
    Private Function Rectangle2Xml(ByVal dic As Dictionary(Of Int32, List(Of KeyValuePair(Of String, String)))) As String

        Dim strFile As String = GetClusterFileName()
        Dim xmlDocument As New XmlDocument
        Dim elemRoot As XmlElement = CreateElement(xmlDocument, "root")
        xmlDocument.AppendChild(elemRoot)
        For Each key1 As KeyValuePair(Of Int32, List(Of KeyValuePair(Of String, String))) In dic
            Dim elemIndex As XmlElement = CreateElement(xmlDocument, "Slide", key1.Key)
            For Each key2 As KeyValuePair(Of String, String) In key1.Value
                Dim eleRectangle As XmlElement = CreateElement(xmlDocument, "Rectangle")
                For Each ele As XmlElement In {CreateElement(xmlDocument, "Item", key2.Key), CreateElement(xmlDocument, "Text", key2.Value.Trim)}
                    eleRectangle.AppendChild(ele)
                Next
                elemIndex.AppendChild(eleRectangle)
            Next
            elemRoot.AppendChild(elemIndex)
        Next

        Dim settings As XmlWriterSettings = New XmlWriterSettings
        With settings
            .Indent = True
            .Encoding = System.Text.Encoding.GetEncoding("Shift_Jis")
            .Async = True
            .NewLineChars = vbCrLf
        End With

        'Dim writer As XmlWriter = XmlWriter.Create(strFile, settings)
        'xmlDocument.Save(writer)
        xmlDocument.Save(strFile)
        Return strFile
    End Function
    ''' <summary>
    ''' グルーピング処理
    ''' </summary>
    ''' <param name="fn">PPTX</param>
    ''' <param name="cluster">クラスタ数</param>
    ''' <param name="lcolumn">スライド上グルーピング処理の列数</param>
    ''' 
    Private Sub Grouping(ByVal fn As String, Optional ByVal cluster As Long = 5, Optional ByVal lcolumn As Long = 3)
        Try

            Do
                If IO.File.Exists(fn) = False Then Exit Do
                Detail(Join({System.Reflection.MethodBase.GetCurrentMethod.Name, fn}, ">"))
                Dim dic1 As Dictionary(Of Int32, List(Of KeyValuePair(Of String, String))) = RdPptx(fn)
                fn = Rectangle2Text(dic1)
                Detail(fn)
                Detail("Call Python...")


                Detail(Join({"Call Pyhton...", fn, cluster}, " "))
                Dim py As String = IO.Path.Combine(System.Environment.CurrentDirectory(), "cluster_ver.0.10.0.py")

                If IO.File.Exists(py) = False Then
                    py = IO.Path.Combine(System.Environment.CurrentDirectory(), "Python\cluster_ver.0.10.0.py")
                End If
                Dim c As New ClassLibraryUtil.clsCluster(fn, cluster, py)

                Me.m_oPowerPoint.shuffle()

                m_dicCluster = c.Clutser(False)
                For Each key1 As KeyValuePair(Of Int32, Dictionary(Of Int32, List(Of ClassLibraryUtil.clsCluster.clsShape))) In m_dicCluster
                    For Each key2 As KeyValuePair(Of Int32, List(Of ClassLibraryUtil.clsCluster.clsShape)) In key1.Value
                        For Each shape As ClassLibraryUtil.clsCluster.clsShape In key2.Value
                            Detail(Join({key1.Key, key2.Key, shape.index, shape.Shape, shape.Text}, ","))
                        Next
                    Next
                Next

                m_oPowerPoint.Align(m_dicCluster, lcolumn)
                Detail(Join({System.Reflection.MethodBase.GetCurrentMethod.Name, "END"}, ":"))
            Loop While (False)
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
        End Try
    End Sub

    ''' <summary>
    ''' ラベリング処理
    ''' </summary>
    ''' <param name="strCategory">MeCab/Janome</param>
    ''' 
    Private Sub Labeling(Optional ByVal strCategory As String = "MeCab")
        Try
            Do
                'クラスタラベリング
                If m_dicCluster.Count = 0 Then Exit Do
                If m_dicLabeling.Count Then m_dicLabeling.Clear()

                For Each key1 As KeyValuePair(Of Int32, Dictionary(Of Int32, List(Of ClassLibraryUtil.clsCluster.clsShape))) In m_dicCluster
                    For Each key2 As KeyValuePair(Of Int32, List(Of ClassLibraryUtil.clsCluster.clsShape)) In key1.Value

                        Dim fn As String = GetClusterFileName("_" & key1.Key & "_" & key2.Key & ".csv")
                        Dim rst As String = GetClusterFileName("_" & key1.Key & "_" & key2.Key & ".rst")

                        Dim lstShape As New List(Of String)
                        For Each s As ClassLibraryUtil.clsCluster.clsShape In key2.Value
                            lstShape.Add(s.Text)
                        Next
                        ClassLibraryUtil.clsUtil.WriteFile(lstShape, fn)

                        Detail("deleting labeling...")
                        m_oPowerPoint.DeleteLabel()

                        Dim py As String = "NLP_VER0.10.0.py"
                        If IO.File.Exists(py) = False Then
                            py = IO.Path.Combine(System.Environment.CurrentDirectory(), "Python\NLP_VER0.10.0.py")
                        End If


                        Detail(Join({"Call Python..."}, " "))
                        Detail(Join({fn, rst, strCategory}, " "))

                        Dim c As New ClassLibraryUtil.clsLabeling(fn, py, strCategory)
                        Dim output As String = c.Labeling(rst).Trim
                        Detail(output)

                        If IO.File.Exists(rst) = False Then Continue For
                        Dim strLabeling As String = Join(ClassLibraryUtil.clsUtil.ReadFile(rst).ToArray, "")
                        If String.IsNullOrEmpty(strLabeling) Then Continue For
                        strLabeling = strLabeling.Trim
                        Detail(strLabeling)

                        'Private m_dicLabeling As New Dictionary(Of Int32, Dictionary(Of Int32, KeyValuePair(Of String, List(Of ClassLibraryUtil.clsCluster.clsShape)))) 
                        If m_dicLabeling.ContainsKey(key1.Key) = False Then
                            m_dicLabeling.Add(key1.Key, New Dictionary(Of Int32, KeyValuePair(Of String, List(Of ClassLibraryUtil.clsCluster.clsShape))))
                        End If

                        If m_dicLabeling(key1.Key).ContainsKey(key2.Key) = False Then
                            m_dicLabeling(key1.Key).Add(key2.Key, New KeyValuePair(Of String, List(Of ClassLibraryUtil.clsCluster.clsShape))(strLabeling, key2.Value))
                        End If
                    Next
                Next
                m_oPowerPoint.Labeling(m_dicLabeling)
                Detail(Join({System.Reflection.MethodBase.GetCurrentMethod.Name, "END"}, ":"))
            Loop While (False)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    ''' <summary>
    ''' グルーピング処理、バックグラウンド実行
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' 
    Private Sub ButtonGrouping_Click(sender As Object, e As EventArgs) Handles ButtonGrouping.Click
        Detail(Join({System.Reflection.MethodBase.GetCurrentMethod.Name, " Grouping..."}, ""))
        Enable(False)
        Me.Opacity = m_Opacity
        Me.StartPosition = FormStartPosition.Manual
        Me.DesktopLocation = New Drawing.Point(0, 0)
        Me.TextBox2.BackColor = Color.Black
        Me.TextBox2.ForeColor = Color.White

        Dim fn As String = Me.TextBox1.Text
        Dim cluster As Long = 5
        Long.TryParse(Me.ComboBox1.Text, cluster)
        If cluster <= 0 Then cluster = 5

        Dim lcolumn As Long = 3
        Long.TryParse(Me.ComboBox2.Text, lcolumn)
        If Not m_frm Is Nothing Then
            m_frm = Nothing
        End If
        m_frm = New ClassLibraryUtil.FormDetail
        m_frm.SetMessage("グルーピング処理中...")
        m_frm.tmr(True)
        Me.BackgroundWorker1.RunWorkerAsync({0, fn, cluster, lcolumn})
    End Sub


    Private Sub FormIIIDEAaaS2Group_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Detail(Join({System.Reflection.MethodBase.GetCurrentMethod.Name, "..."}, ""))
        ComboBox3.SelectedIndex = 0
        Enable(False)
        Enable(True, {Me.Button1})
        Me.Button1.Select()
    End Sub

    ''' <summary>
    ''' ラベリング処理、バックグラウンド実行
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' 
    Private Sub ButtonLabeling_Click(sender As Object, e As EventArgs) Handles ButtonLabeling.Click
        Detail(Join({System.Reflection.MethodBase.GetCurrentMethod.Name, "Labeling..."}, ">"))
        Enable(False)
        Me.Opacity = m_Opacity
        Me.StartPosition = FormStartPosition.Manual
        Me.DesktopLocation = New Drawing.Point(0, 0)
        Me.TextBox2.BackColor = Color.Black
        Me.TextBox2.ForeColor = Color.White
        If Not m_frm Is Nothing Then
            m_frm = Nothing
        End If
        m_frm = New ClassLibraryUtil.FormDetail
        m_frm.SetMessage("ラベリング処理中...")
        m_frm.tmr(True)
        Dim strCategory As String = Me.ComboBox3.Text
        Me.BackgroundWorker1.RunWorkerAsync({1, strCategory})
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Dim fn As String = CType(sender, TextBox).Text
        If String.IsNullOrEmpty(fn) OrElse IO.File.Exists(fn) = False Then
            Enable(False)
        Else
            Enable(True)
        End If
    End Sub

    ''' <summary>
    ''' バックグラウンド処理：グルーピング/ラベリング
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' 
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim lCat As Long = CInt(e.Argument(0))
        If lCat = 0 Then
            'グルーピング
            Dim fn As String = CStr(e.Argument(1))
            Grouping(fn, CInt(e.Argument(2)), CInt(e.Argument(3)))
        Else
            'ラベリング
            Labeling(CStr(e.Argument(1)))
        End If
    End Sub

    ''' <summary>
    ''' バックグラウンド完了通知、画面活性化
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' 
    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Me.Opacity = 1
        Enable(True)
        Me.TextBox2.BackColor = SystemColors.Control
        Me.TextBox2.ForeColor = SystemColors.WindowText


        If Not m_frm Is Nothing Then
            m_frm.tmr(False)
            m_frm.Dispose()
            m_frm = Nothing
        End If

        Me.Activate()
        Me.TopMost = True
        Me.TopMost = False

    End Sub

    Private Sub FormIIIDEAaaS2Group_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MyBase.KeyPress
        e.Handled = True
    End Sub
End Class

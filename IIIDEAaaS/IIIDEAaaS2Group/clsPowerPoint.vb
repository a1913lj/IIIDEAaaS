Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.PowerPoint
''' <summary>
''' POWERPOINT処理クラス
''' </summary>
''' 
Public Class clsPowerPoint
    Implements IDisposable

    Protected _disposed = False
    Dim m_oApp As Microsoft.Office.Interop.PowerPoint.Application = Nothing
    Dim m_oPres As PowerPoint.Presentations = Nothing
    Dim m_oPre As PowerPoint.Presentation = Nothing

    Dim m_lstLabelingShape As New List(Of PowerPoint.Shape)
    Dim m_lstLineFormat As New List(Of PowerPoint.Shape)

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
                MRC({m_oPre, m_oPres, m_oApp})
                m_oApp = New PowerPoint.Application
                MRC(m_oApp)
            End If
            '*** マネージドリソースの開放

        End If
        _disposed = True
    End Sub

    Public Sub New()
        m_oApp = New Microsoft.Office.Interop.PowerPoint.Application
        With m_oApp
            m_oPres = .Presentations
        End With
    End Sub
    ''' <summary>
    ''' アンマネージリソースの開放
    ''' </summary>
    ''' <param name="obj">オブジェクト</param>
    ''' 
    Private Sub MRC(obj As Object)
        If obj IsNot Nothing Then
            Try
                If obj Is m_oPre Then CType(obj, PowerPoint.Presentation).Close()
                If obj Is m_oApp Then CType(obj, Microsoft.Office.Interop.PowerPoint.Application).Quit()
            Catch ex As Exception
            Finally
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
                obj = Nothing
            End Try
            System.GC.Collect()
        End If
        obj = Nothing
    End Sub
    ''' <summary>
    ''' アンマネージリソースの開放
    ''' </summary>
    ''' <param name="objLst">オブジェクト</param>
    ''' 
    Private Sub MRC(objLst As List(Of Object))
        For Each obj In objLst
            MRC(obj)
        Next
    End Sub
    ''' <summary>
    ''' アンマネージリソースの開放
    ''' </summary>
    ''' <param name="objs">オブジェクト</param>
    ''' 
    Private Sub MRC(objs() As Object)
        For Each obj In objs
            MRC(obj)
        Next
    End Sub
    ''' <summary>
    ''' POWERPOINT処理、SHAPE一覧取得（テキスト関連のみ）
    ''' </summary>
    ''' <param name="oPre">スライド</param>
    ''' <returns>SHAPE一覧を辞書型返す</returns>
    ''' 
    Public Function RdPptx(oPre As PowerPoint.Presentation) As Dictionary(Of Int32, List(Of KeyValuePair(Of String, String)))
        Me.Delete(Me.m_lstLabelingShape)
        Dim dic As New Dictionary(Of Int32, List(Of KeyValuePair(Of String, String)))
        Try
            For Each slide As PowerPoint.Slide In oPre.Slides
                Dim lstRet As New List(Of KeyValuePair(Of String, String))
                For Each shape As PowerPoint.Shape In slide.Shapes
                    '値	定数	定義
                    '-2	msoShapeTypeMixed	
                    '1   msoAutoShape	図形・オートシェイプ
                    '2   msoCallout	吹き出し
                    '3   msoChart	グラフ
                    '4   msoComment	コメント
                    '5   msoFreeform	フリーフォーム
                    '6   msoGroup	グループ化された図形
                    '7   msoEmbeddedOLEObject	埋め込みOLEオブジェクト
                    '8   msoFormControl	フォームコントロール
                    '9   msoLine	線
                    '10  msoLinkedOLEObject	リンクOLEオブジェクト
                    '11  msoLinkedPicture	リンク画像
                    '12  msoOLEControlObject	ActiveXコントロール
                    '13  msoPicture	画像
                    '14  msoPlaceholder	プレースホルダー
                    '15  msoTextEffect	テキスト効果
                    '16  msoMedia	メディア
                    '17  msoTextBox	テキストボックス
                    '18  msoScriptAnchor	スクリプトアンカー
                    '19  msoTable	表
                    '20  msoCanvas	描画キャンバス
                    '21  msoDiagram	図表
                    '22  msoInk	インク
                    '23  msoInkComment	インクコメント
                    '24  msoSmartArt	スマートアート
                    '25  msoSlicer	スライサー
                    '26  msoWebVideo	Webビデオ
                    '27  msoContentApp	コンテンツアドイン
                    '28  msoGraphic	グラフィック
                    '29  msoLinkedGraphic	リンクグラフィック
                    '30  mso3DModel	3Dモデル
                    '31  msoLinked3DModel	リンク3Dモデル
                    If (shape.Type = MsoShapeType.msoAutoShape OrElse shape.Type = MsoShapeType.msoTextBox) AndAlso shape.TextFrame.HasText Then
                        Dim strText As String = shape.TextFrame.TextRange.Text
                        If String.IsNullOrEmpty(strText) Then Continue For
                        lstRet.Add(New KeyValuePair(Of String, String)(shape.Name, strText))
                    End If
                Next
                If lstRet.Count = 0 Then Continue For
                If dic.ContainsKey(slide.SlideIndex) = False Then
                    dic.Add(slide.SlideIndex, New List(Of KeyValuePair(Of String, String)))
                End If
                dic(slide.SlideIndex) = lstRet
            Next
        Catch ex As Exception
        Finally
        End Try
        Return dic
    End Function

    ''' <summary>
    ''' POWERPOINT読み込み処理
    ''' </summary>
    ''' <param name="fn">ファイル名</param>
    ''' <returns>true/false</returns>
    ''' 
    Public Function RdOpen(fn As String) As Boolean
        Try
            If m_oPre IsNot Nothing Then
                MRC(m_oPre)
            End If
            m_oPre = m_oPres.Open(fn, MsoTriState.msoTrue)
            m_oPre.Application.WindowState = PpWindowState.ppWindowMaximized
            m_oPre.Application.Activate()
        Catch ex As Exception
            MRC(m_oPre)
        Finally
        End Try
        Return m_oPre IsNot Nothing
    End Function

    ''' <summary>
    ''' POWERPOINT読み込み処理
    ''' </summary>
    ''' <param name="fn">ファイル名</param>
    ''' <returns>SHAPE一覧を辞書型返す</returns>
    ''' 
    Public Function RdPptx(fn As String) As Dictionary(Of Int32, List(Of KeyValuePair(Of String, String)))
        Dim dic As New Dictionary(Of Int32, List(Of KeyValuePair(Of String, String)))
        Do
            If m_oPre Is Nothing Then RdOpen(fn)
            m_oPre.Application.WindowState = PpWindowState.ppWindowMaximized
            dic = RdPptx(m_oPre)
        Loop While (False)

        Return dic
    End Function

    ''' <summary>
    ''' shapeごちゃ混ぜ
    ''' </summary>
    ''' <returns>0/1:OK/NG</returns>
    Public Function shuffle() As Long
        Dim lret As Long = 0
        Try
            Dim w = m_oPre.PageSetup.SlideWidth / 10 * 5
            Dim h = m_oPre.PageSetup.SlideHeight / 10 * 5
            For Each sld As PowerPoint.Slide In m_oPre.Slides
                For Each shape As PowerPoint.Shape In sld.Shapes
                    'If shape.HasTextFrame = MsoTriState.msoFalse Then Continue For
                    'Dim strText As String = shape.TextFrame.TextRange.Text
                    'If String.IsNullOrEmpty(strText) Then Continue For
                    Randomize(System.Environment.TickCount)
                    Dim rd As New Random(System.Environment.TickCount)
                    System.Threading.Thread.Sleep(1)

                    Dim w1 As Single = rd.Next(0, w)
                    Dim h1 As Single = rd.Next(0, h)
                    With shape
                        .Left = w1
                        .Top = h1
                        .ZOrder(MsoZOrderCmd.msoBringToFront)
                    End With

                Next
            Next
        Catch ex As Exception

        End Try

        Return lret
    End Function


    Private Function AddLine(sld As PowerPoint.Slide, ByVal xyBegin As KeyValuePair(Of Single, Single), ByVal xyEnd As KeyValuePair(Of Single, Single)) As PowerPoint.Shape
        'Object.AddLine(BeginX, BeginY, EndX, EndY)
        'BeginX      左上隅を基点とした線の始点のX座標［省略不可］
        'BeginY      左上隅を基点とした線の始点のY座標［省略不可］
        'EndX        左上隅を基点とした線の終点のX座標［省略不可］
        'EndY        左上隅を基点とした線の終点のY座標［省略不可］
        Dim shp As PowerPoint.Shape = sld.Shapes.AddLine(xyBegin.Key, xyBegin.Value, xyEnd.Key, xyEnd.Value)
        With shp.Line
            .DashStyle = MsoLineDashStyle.msoLineDashDotDot
            .ForeColor.RGB = RGB(50, 0, 128)
            .BeginArrowheadLength = MsoArrowheadLength.msoArrowheadShort
            .BeginArrowheadStyle = MsoArrowheadStyle.msoArrowheadOval
            .BeginArrowheadWidth = MsoArrowheadWidth.msoArrowheadNarrow
            .EndArrowheadLength = MsoArrowheadLength.msoArrowheadLong
            .EndArrowheadStyle = MsoArrowheadStyle.msoArrowheadTriangle
            .EndArrowheadWidth = MsoArrowheadWidth.msoArrowheadWide
        End With
        Return shp
    End Function

    ''' <summary>
    ''' グルーピング
    ''' </summary>
    ''' <param name="sld">スライド</param>
    ''' <param name="dic">シェイプ一覧</param>
    ''' <param name="c">画面上コラム数</param>
    ''' <param name="lWait">待ち時間</param>
    ''' <returns></returns>
    ''' 
    Private Function Align(sld As PowerPoint.Slide, dic As Dictionary(Of Int32, List(Of ClassLibraryUtil.clsCluster.clsShape)), Optional ByVal c As Int32 = 3, Optional ByVal lWait As Long = 100) As Long
        Dim lRet As Long = 0
        Do
            Dim w = m_oPre.PageSetup.SlideWidth
            Dim h = m_oPre.PageSetup.SlideHeight

            Dim dicShape As New Dictionary(Of String, PowerPoint.Shape)
            For Each shape As PowerPoint.Shape In sld.Shapes
                If shape.HasTextFrame = MsoTriState.msoFalse Then Continue For
                Dim strText As String = shape.TextFrame.TextRange.Text
                If String.IsNullOrEmpty(strText) Then Continue For
                dicShape.Add(shape.Name, shape)
            Next


            Dim dic1 As New Dictionary(Of Int32, List(Of KeyValuePair(Of PowerPoint.Shape, ClassLibraryUtil.clsCluster.clsShape)))
            For Each k As KeyValuePair(Of String, PowerPoint.Shape) In dicShape
                k.Value.Left = -1000
                k.Value.Top = -1000
                Threading.Thread.Sleep(1)
            Next

            For Each key As KeyValuePair(Of Int32, List(Of ClassLibraryUtil.clsCluster.clsShape)) In dic
                For Each shape In key.Value
                    If dicShape.ContainsKey(shape.Shape) Then
                        If dic1.ContainsKey(key.Key) = False Then
                            dic1.Add(key.Key, New List(Of KeyValuePair(Of PowerPoint.Shape, ClassLibraryUtil.clsCluster.clsShape)))
                        End If
                        shape.Width = dicShape(shape.Shape).Width
                        shape.Height = dicShape(shape.Shape).Height
                        dic1(key.Key).Add(New KeyValuePair(Of PowerPoint.Shape, ClassLibraryUtil.clsCluster.clsShape)(dicShape(shape.Shape), shape))
                    End If
                Next
            Next

            Dim woffset As Single = w / (10 + 2) '左アライメント調整
            Dim hoffset As Single = h / (10 + 2) 'トップアライメント調整
            w -= woffset * 2
            h -= hoffset * 2

            '並べ替え
            Dim r As Long = Math.Ceiling(dic1.Count / c)
            Dim xy(r - 1, c - 1) As Single
            For Each key In dic1
                Dim r1 As Long = Int(key.Key / c)
                Dim c1 As Long = key.Key Mod c
                Dim w1 As Single = woffset + (w / c) * c1
                Dim h1 As Single = hoffset + (h / r) * r1

                Dim i As Int32 = 0
                Dim refLeft As Single = w1
                Dim refTop As Single = h1

                Dim Height As Single = 0#
                For Each s In key.Value
                    Height += s.Value.Height
                Next

                For Each s In key.Value
                    With s.Key
                        .Left = refLeft
                        .Top = refTop
                        .ZOrder(MsoZOrderCmd.msoBringToFront)
                        i += 1

                        If Height <= (h / r) Then
                            refLeft = .Left
                            refTop = .Top + .Height
                        Else
                            refLeft += (.Width / 20)
                            refTop = h1 + (h / r / key.Value.Count) * (i)
                        End If
                    End With
                    System.Threading.Thread.Sleep(lWait)
                Next
            Next
        Loop While (False)
        Return lRet
    End Function
    Public Function Align(dic As Dictionary(Of Int32, Dictionary(Of Int32, List(Of ClassLibraryUtil.clsCluster.clsShape))), Optional ByVal column As Long = 3) As Long
        Dim lRet As Long = 0
        Do
            If m_oPre Is Nothing Then Exit Do
            'm_oPre.Application.Activate()
            m_oPre.Application.WindowState = PpWindowState.ppWindowMaximized
            For Each sld As PowerPoint.Slide In m_oPre.Slides
                For Each key As KeyValuePair(Of Int32, Dictionary(Of Int32, List(Of ClassLibraryUtil.clsCluster.clsShape))) In dic
                    If sld.SlideIndex = key.Key Then
                        Align(sld, key.Value, column)
                    End If
                Next
            Next
        Loop While (False)
        Return lRet
    End Function

    'グループ、レベリング, シェイプ 
    ''' <summary>
    ''' ラベリング処理
    ''' </summary>
    ''' <param name="sld">スライド</param>
    ''' <param name="shape">シェイプ</param>
    ''' <param name="strText">ラベリング結果</param>
    ''' <returns>ラベリングのシェイプ情報作成</returns>
    ''' 
    Private Function Labeling(sld As PowerPoint.Slide, shape As PowerPoint.Shape, Optional ByVal strText As String = "Test text") As PowerPoint.Shape
        Dim sr As PowerPoint.Shape = Nothing
        Do
            'オートシェイプを作成します。 新しいオートシェイプを表す**Shape** オブジェクトを返します。
            'https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
            sr = sld.Shapes.AddShape(Type:=MsoAutoShapeType.msoShapeRectangularCallout,'msoShapeOval,
                                     Left:=shape.Left, Top:=shape.Top, Width:=shape.Width, Height:=shape.Height)
            With sr.TextFrame
                .TextRange.Text = strText
                With .TextRange.Font
                    .Name = "Meiryo UI"
                    .NameFarEast = "Meiryo UI"
                    .Size = 10.5
                End With
                .Parent.Rotation = 45
                If .TextRange.Characters.Count < 50 Then

                End If
                .AutoSize = PpAutoSize.ppAutoSizeShapeToFitText

            End With

        Loop While (False)
        Return sr
    End Function
    ''' <summary>
    ''' ラベリング処理
    ''' </summary>
    ''' <param name="sld">スライド</param>
    ''' <param name="dic">シェイプ情報</param>
    ''' <returns>ラベリング一覧辞書情報</returns>
    ''' 
    Private Function Labeling(sld As PowerPoint.Slide, dic As Dictionary(Of Int32, KeyValuePair(Of String, List(Of ClassLibraryUtil.clsCluster.clsShape)))) As List(Of PowerPoint.Shape)
        Dim lstShape As New List(Of PowerPoint.Shape)
        Do
            Dim dicShape As New Dictionary(Of String, PowerPoint.Shape)
            For Each shape As PowerPoint.Shape In sld.Shapes
                If shape.HasTextFrame = MsoTriState.msoFalse Then Continue For
                Dim strText As String = shape.TextFrame.TextRange.Text
                If String.IsNullOrEmpty(strText) Then Continue For
                dicShape.Add(shape.Name, shape)
            Next

            For Each key In dic
                For Each s In key.Value.Value
                    For Each shape In dicShape
                        If s.Shape <> shape.Key Then Continue For
                        lstShape.Add(Labeling(sld, shape.Value, key.Value.Key))
                        Exit For
                    Next
                    Exit For
                Next
            Next

        Loop While (False)
        Return lstShape
    End Function
    ''' <summary>
    ''' ラベリング処理
    ''' </summary>
    ''' <param name="dic">ラベリング情報</param>
    ''' <param name="lWait">待ち時間</param>
    ''' <returns>ラベリング一覧辞書情報</returns>
    ''' 
    Public Function Labeling(dic As Dictionary(Of Int32, Dictionary(Of Int32, KeyValuePair(Of String, List(Of ClassLibraryUtil.clsCluster.clsShape)))), Optional ByVal lWait As Long = 1000) As List(Of PowerPoint.Shape)
        Dim lstShape As New List(Of PowerPoint.Shape)
        Do
            If m_oPre Is Nothing Then Exit Do
            'm_oPre.Application.Activate()
            m_oPre.Application.WindowState = PpWindowState.ppWindowMaximized
            DeleteLabel()
            For Each sld As PowerPoint.Slide In m_oPre.Slides
                For Each key In dic
                    If sld.SlideIndex = key.Key Then
                        lstShape.AddRange(Labeling(sld, key.Value))
                        System.Threading.Thread.Sleep(lWait)
                    End If
                Next
            Next
        Loop While (False)
        m_lstLabelingShape = lstShape
        Return lstShape
    End Function

    ''' <summary>
    ''' シェイプ削除
    ''' </summary>
    ''' <param name="objs">オブジェクト（シェイプ情報）</param>
    ''' 
    Private Sub Delete(objs As List(Of PowerPoint.Shape))
        Try
            For Each obj In objs
                Try
                    obj.Delete()
                Catch ex As Exception
                End Try
            Next
            objs.Clear()
        Catch ex As Exception
        End Try
    End Sub
    ''' <summary>
    ''' ラベリング削除
    ''' </summary>
    ''' 
    Public Sub DeleteLabel()
        Delete(m_lstLabelingShape)
    End Sub
End Class

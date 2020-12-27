Imports System.Windows.Forms.DataVisualization.Charting

''' <summary>
''' チーム分け処理メイン画面
''' </summary>
Public Class FormIIIDEAaaSTeams

    Dim _lst As New List(Of clsHerrmann)
    Dim _dic As New Dictionary(Of String, clsTeam)

    Dim _frm As New ClassLibraryUtil.FormDetail

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    End Sub

    Private Function OpenFileDialog() As String
        Dim strFile As String = ""
        Do
            Dim ofd As New OpenFileDialog
            With ofd
                .Filter = "ハーマン・モデル結果ファイル(*.csv)|*.csv|すべてのファイル(*.*)|*.*"
                .Title = "ハーマン・モデルの集計結果ファイルを選択してください"
                .FilterIndex = 1
                .CheckFileExists = True
            End With
            If ofd.ShowDialog <> Windows.Forms.DialogResult.OK Then Exit Do
            strFile = ofd.FileName
            _lst.Clear()
            _dic.Clear()
            For Each dgv As DataGridView In {Me.DataGridView1, Me.DataGridView2}
                dgv.Rows.Clear()
            Next
            Me.ComboBox2.Items.Clear()
            Me.ComboBox3.Items.Clear()
            Chart()
        Loop While (False)

        Return strFile
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim fn As String = OpenFileDialog()
        Me.TextBox1.Text = fn
        _lst = ReadHerrman(fn)
        Herrman2DataGridView(_lst, Me.DataGridView1)
        Me.Button2.Enabled = _lst.Count <> 0
    End Sub

    Private Sub Herrman2DataGridView(ByVal lst As List(Of clsHerrmann), sender As Object)
        Dim dgv As DataGridView = DirectCast(sender, DataGridView)
        dgv.Rows.Clear()
        For Each c As clsHerrmann In lst
            dgv.Rows.Add({c.N, c.A, c.B, c.C, c.D})
        Next
    End Sub

    Private Sub Herrman2DataGridView(ByVal dic As Dictionary(Of String, clsTeam), sender As Object)
        Dim dgv As DataGridView = DirectCast(sender, DataGridView)
        dgv.Rows.Clear()
        For Each k As KeyValuePair(Of String, clsTeam) In dic
            Dim b As Boolean = True
            For Each c As clsHerrmann In k.Value.Menbers
                dgv.Rows.Add({IIf(b, k.Key, ""), c.N, c.A, c.B, c.C, c.D})
                b = False
            Next
        Next
    End Sub

    Private Sub Herrman2Combox(ByVal dic As Dictionary(Of String, clsTeam), sender As Object)
        Dim obj As ComboBox = DirectCast(sender, ComboBox)
        obj.Items.Clear()
        For Each k As KeyValuePair(Of String, clsTeam) In dic
            obj.Items.Add(k.Key)
        Next
        If obj.Items.Count Then obj.SelectedIndex = 0
    End Sub

    Private Sub Herrman2Combox(ByVal strGrp As String, sender As Object)
        Herrman2Combox(_dic, strGrp, sender)
    End Sub

    Private Sub Herrman2Combox(ByVal dic As Dictionary(Of String, clsTeam), ByVal strGrp As String, sender As Object)
        Dim obj As ComboBox = DirectCast(sender, ComboBox)
        obj.Items.Clear()
        If dic.ContainsKey(strGrp) Then
            Herrman2Combox(dic(strGrp), sender)
        End If
    End Sub
    Private Sub Herrman2Combox(ByVal Team As clsTeam, sender As Object)
        Dim obj As ComboBox = DirectCast(sender, ComboBox)
        obj.Items.Clear()
        For Each c As clsHerrmann In Team.Menbers
            obj.Items.Add(c.N)
        Next
        If obj.Items.Count Then obj.SelectedIndex = 0
    End Sub

    ''' <summary>
    ''' ハンマーモデル結果
    ''' </summary>
    ''' <param name="fn">ハンマーモデル結果</param>
    ''' <returns>メンバー個々情報</returns>
    ''' 
    Private Function ReadHerrman(ByVal fn As String) As List(Of clsHerrmann)
        Dim lstHerrman As New List(Of clsHerrmann)
        If IO.File.Exists(fn) Then
            Dim lst As List(Of String) = ClassLibraryUtil.clsUtil.ReadFile(fn, System.Text.Encoding.GetEncoding("Shift-JIS"))
            For Each s As String In lst
                Dim ss As String() = Split(s, ",")
                If ss.Length < 6 Then Continue For
                Do
                    Dim c As New clsHerrmann("")
                    c.N = ss(1)
                    For idx As Int32 = 2 To 5
                        Dim l As Long = 0
                        If Long.TryParse(ss(idx), l) = False Then Exit Do
                        If idx = 2 Then c.A = l
                        If idx = 3 Then c.B = l
                        If idx = 4 Then c.B = l
                        If idx = 5 Then c.D = l
                    Next
                    lstHerrman.Add(c)
                Loop While (False)
            Next
        End If
        Return lstHerrman
    End Function


    Private Sub ChartClar(ByVal cht As System.Windows.Forms.DataVisualization.Charting.Chart)
        'Chart の設定を初期値に戻す(通常は必要ありません)
        'With cht
        '    .Titles.Clear()                  'タイトルの初期化
        '    .BackGradientStyle = GradientStyle.None
        '    .BackColor = Color.White         '背景色を白色に
        '    .BorderColor = .BackColor
        '    '外形をデフォルトに
        '    .BorderSkin.SkinStyle = BorderSkinStyle.None
        '    .Legends.Clear()                 '凡例の初期化
        '    .Legends.Add("Legend1")
        '    .Series.Clear()                  '系列(データ関係)の初期化
        '    .ChartAreas.Clear()              '軸メモリ・3D 表示関係の初期化
        '    .ChartAreas.Add("ChartArea1")
        '    .Annotations.Clear()
        'End With
    End Sub
    ''' <summary>
    ''' レーターチャット用レコード取得
    ''' </summary>
    ''' <param name="strGrp"></param>
    ''' <param name="strUser"></param>
    ''' <returns>DataSet情報</returns>
    ''' 
    Private Function GetData(ByVal strGrp As String, ByVal strUser As String) As DataSet
        Dim ds As New DataSet
        Dim dt As New DataTable

        For i As Int32 = 1 To 5
            dt.Columns.Add(DataGridView2.Columns(i).HeaderText)
        Next
        ds.Tables.Add(dt)

        Do
            If _dic.ContainsKey(strGrp) = False Then Exit Do
            If _dic(strGrp).Menbers.Count = 0 Then Exit Do

            For Each c As clsHerrmann In _dic(strGrp).Menbers

                If strUser = "" Then
                    Dim dtRow As DataRow = ds.Tables(0).NewRow()
                    For i As Int32 = 0 To 4
                        dtRow(i) = {c.N, c.A, c.B, c.C, c.D}(i)
                    Next
                    ds.Tables(0).Rows.Add(dtRow)
                Else
                    If strUser = c.N Then
                        Dim dtRow As DataRow = ds.Tables(0).NewRow()
                        For i As Int32 = 0 To 4
                            dtRow(i) = {c.N, c.A, c.B, c.C, c.D}(i)
                        Next
                        ds.Tables(0).Rows.Add(dtRow)
                        Exit For
                    End If
                End If
            Next
        Loop While (False)
        Return ds
    End Function
    ''' <summary>
    ''' レーターチャット描画
    ''' </summary>
    ''' <param name="bSave">画像ファイルとして保存有無フラグ</param>
    ''' 
    Private Sub Chart(Optional ByVal bSave As Boolean = False)
        Dim strGrp As String = Me.ComboBox2.Text
        Dim struser As String = IIf(CheckBox1.Checked, "", Me.ComboBox3.Text)
        Dim ds As DataSet = GetData(strGrp, struser)

        If String.IsNullOrEmpty(strGrp) = False AndAlso String.IsNullOrEmpty(struser) = False Then
            With Chart1
                .DataSource = ds
                .Series.Clear()
                .Titles.Clear()
                .Legends(0).Docking = Docking.Top
                .Legends(0).LegendStyle = LegendStyle.Table
                .Legends(0).Enabled = True

                .Series.Add(struser)
                For Each c As clsHerrmann In _dic(strGrp).Menbers
                    If c.N = struser Then
                        Dim yValues() As Long = {c.A, c.B, c.C, c.D}
                        Dim xValues() As String = {c.AN, c.BN, c.CN, c.DN}
                        .Series(struser).Points.DataBindXY(xValues, yValues)
                        .Series(struser).ChartType = SeriesChartType.Radar
                        '.Series(struser)("AreaDrawingStyle") = "Polygon"
                        '.Series(struser)("CircularLabelsStyle") = "Horizontal"
                        Exit For
                    End If
                Next
            End With
            Return
        End If

        With Chart1
            .DataSource = ds
            .Series.Clear()
            .Titles.Clear()
            .Legends(0).Docking = Docking.Top
            .Legends(0).LegendStyle = LegendStyle.Table
            .Legends(0).Enabled = True

            For i As Integer = 1 To ds.Tables(0).Columns.Count - 1
                Dim columnName As String = ds.Tables(0).Columns(i).ColumnName.ToString
                .Series.Add(columnName)
                .Series(columnName).ChartType = DataVisualization.Charting.SeriesChartType.Column
                .Series(columnName).XValueMember = ds.Tables(0).Columns(0).ColumnName.ToString()
                .Series(columnName).YValueMembers = columnName
                '.Series(columnName).LegendText = columnName.Substring(0, 1)
                .Series(columnName).LegendText = columnName.ToString
            Next

            For i As Integer = 0 To .Series.Count - 1
                If struser = "" Then
                    .Series(i).ChartType = DataVisualization.Charting.SeriesChartType.Radar
                End If
            Next

            Dim Col() As Color = {Color.DeepSkyBlue, Color.DeepPink, Color.ForestGreen, Color.Orange}
            For i As Integer = 0 To ds.Tables(0).Columns.Count - 1 - 1
                'グラフの色を透明に設定(重なった場合後ろが見えないので)
                .Series(i).Color = Color.Transparent
                'グラフの線の色を設定
                .Series(i).BorderColor = Col(i)
                '線の太さを設定
                .Series(i).BorderWidth = 3
            Next i


            With .ChartAreas(0)
                With .AxisX
                    .LabelAutoFitMaxFontSize = 10
                    Dim i As Int32 = 0
                    'For Each c As clsHerrmann In _dic(strGrp)
                    '    .CustomLabels.Add(i * 2, 0, c.N & IIf(lst.Max = c.D, ":Facilitator", ""))
                    'Next
                    '.IsLabelAutoFit = True
                    '.LabelAutoFitStyle = LabelAutoFitStyles.DecreaseFont Or LabelAutoFitStyles.IncreaseFont
                End With


                With .AxisY
                    .Maximum = 100    '点数の最大値
                    .Minimum = 0     '点数の最小値
                    .Interval = 20    '点数のメモリ間隔(２０点毎)
                    .LabelStyle.ForeColor = Color.Red
                    .MajorGrid.LineWidth = 1

                    '.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                    ''10点毎に補助線を表示
                    '.MinorGrid.Enabled = True   'True に設定しないと表示しない
                    '.MinorGrid.Interval = 10
                    '.MinorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                End With
            End With

            If bSave Then
                Dim strImagePath As String = IO.Path.Combine(Environment.CurrentDirectory, "image")
                If IO.Directory.Exists(strImagePath) = False Then
                    IO.Directory.CreateDirectory(strImagePath)
                End If
                Dim fn As String = IO.Path.Combine(strImagePath, strGrp & ".png")
                If IO.File.Exists(fn) = True Then
                    IO.File.SetAttributes(fn, IO.FileAttributes.Normal)
                    IO.File.Delete(fn)
                End If
                .SaveImage(fn, System.Drawing.Imaging.ImageFormat.Png)
            End If
        End With

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        Dim cb As CheckBox = DirectCast(sender, CheckBox)

        If cb.Checked Then
            cb.Text = "チーム"
            ComboBox3.Visible = False
        Else
            cb.Text = "個人"
            ComboBox3.Visible = True
        End If
        Chart()
        ChangeDgvSelect()
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim fn As String = e.Argument(0)
        Dim lTeamNumber As Long = e.Argument(1)
        Do
            _dic.Clear()
            If IO.File.Exists(fn) = False Then Exit Do
            Dim strPath As String = IO.Path.GetDirectoryName(fn)
            Dim rst As String = IO.Path.Combine(strPath, IO.Path.GetFileNameWithoutExtension(fn) & "_Result.csv")
            If IO.File.Exists(rst) Then
                IO.File.SetAttributes(rst, IO.FileAttributes.Normal)
                IO.File.Delete(rst)
            End If

            For Each pathFrom As String In IO.Directory.EnumerateFiles(strPath, "CI_*.csv")
                IO.File.SetAttributes(pathFrom, IO.FileAttributes.Normal)
                IO.File.Delete(pathFrom)
            Next

            'Python チーム分け
            Dim pst As String = IO.Path.Combine(System.Environment.CurrentDirectory, "iiidea_team.bat")
            Dim psi As New System.Diagnostics.ProcessStartInfo()
            psi.FileName = pst
            psi.Arguments = """" & strPath & """" & " " & """" & fn & """" & " " & """" & rst & """" & " " & lTeamNumber.ToString
            psi.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden
            'アプリケーションを起動する
            Dim p As System.Diagnostics.Process = System.Diagnostics.Process.Start(psi)
            p.Start()
            p.WaitForExit()

            If IO.File.Exists(rst) = False Then Exit Do
            Dim strGrp As String = ""
            Dim lst As List(Of String) = ClassLibraryUtil.clsUtil.ReadFile(rst)
            If lst.Count Then lst.RemoveAt(0)

            For Each s As String In lst
                Dim ss As String() = Split(s, ",")
                If ss.Length < 7 Then Continue For

                If String.IsNullOrEmpty(ss(0)) = False Then
                    strGrp = "チーム" & ss(0)
                End If
                Do

                    Dim c As New clsHerrmann("")
                    c.N = ss(2)
                    For idx As Int32 = 3 To 6
                        Dim l As Long = 0
                        If Long.TryParse(ss(idx), l) = False Then Exit Do
                        If idx = 3 Then c.A = l
                        If idx = 4 Then c.B = l
                        If idx = 5 Then c.C = l
                        If idx = 6 Then c.D = l
                    Next

                    If _dic.ContainsKey(strGrp) = False Then
                        _dic.Add(strGrp, New clsTeam(strGrp))
                    End If

                    For Each t As clsHerrmann In _dic(strGrp).Menbers
                        If String.Compare(c.N, t.N, True) = 0 Then
                            Exit Do
                        End If
                    Next
                    _dic(strGrp).Menbers.Add(c)
                Loop While (False)
            Next
        Loop While (False)
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Dim btn As Button = Me.Button2

        Me.CheckBox1.Checked = True
        Me.TabControl2.SelectedIndex = 1
        Herrman2DataGridView(_dic, DataGridView2)
        Herrman2Combox(_dic, ComboBox2)

        Dim i As Long = 0
        For Each s As String In _dic.Keys
            Me.ComboBox2.SelectedIndex = i
            Chart(True)
            i += 1
        Next

        MakeTeamHtml(Me.CheckBox2.Checked)
        If Me.ComboBox2.Items.Count Then
            Me.ComboBox2.SelectedIndex = 0
            Chart()
        End If

        btn.Enabled = True
        If Me.CheckBox2.Checked = False Then
            Me.Activate()
            Me.TopMost = True
            Me.TopMost = False
        End If

        _frm.tmr(False)
    End Sub
    ''' <summary>
    ''' チーム情報基、チーム分け＆レーターチャット描画
    ''' 処理中表示する為、バックグラウンド処理変更
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim lTeamNumber As Long = 5
        Long.TryParse(Me.CheckBox1.Text, lTeamNumber)
        If lTeamNumber <= 0 Then lTeamNumber = 5

        Dim fn As String = Me.TextBox1.Text
        CType(sender, Button).Enabled = False
        _frm.tmr(True)
        Me.BackgroundWorker1.RunWorkerAsync({fn, lTeamNumber})
    End Sub

    ''' <summary>
    ''' レーターチャットの結果をHTML表示
    ''' </summary>
    ''' <param name="bShow">HTML表示有無</param>
    ''' 
    Private Sub MakeTeamHtml(Optional ByVal bShow As Boolean = False)
        Do
            Dim tmp As String = IO.Path.Combine(System.Environment.CurrentDirectory, "iiidea_result.tmp")
            Dim img As String = IO.Path.Combine(Environment.CurrentDirectory, "image")
            Dim htm As String = IO.Path.Combine(Environment.CurrentDirectory, IO.Path.GetFileNameWithoutExtension(tmp) & ".html")
            '<table>
            '    <tr>
            '        <th>チーム</th>
            '        <th>レーダーチャート</th>
            '        <th>チーム</th>
            '        <th>レーダーチャート</th>
            '        <th>チーム</th>
            '        <th>レーダーチャート</th>
            '    </tr>
            '    <tr>
            '        <td><center>チーム0<br><font color="blue"><u>AAA</u></font></center></td>
            '        <td><img src = "./image/チーム0.png" ></td>
            '        <td>チーム1</td>
            '        <td><img src = "./image/チーム1.png" ></td>
            '        <td>チーム2</td>
            '        <td><img src = "./image/チーム2.png" ></td>
            '    </tr>
            '    <tr>
            '        <td>チーム3</td>
            '        <td><img src = "./image/チーム3.png" ></td>
            '        <td>チーム4</td>
            '        <td><img src = "./image/チーム4.png" ></td>
            '        <td></td>
            '        <td></td>
            '    </tr>
            '</table>

            If IO.File.Exists(tmp) = False Then Exit Do
            If _dic.Count = 0 Then Exit Do

            Dim lst As New List(Of String)
            lst.Add("<table>")
            lst.Add("<tr>")
            For i As Int32 = 0 To 2
                For Each s As String In {"<th>チーム</th>", "<th>レーダーチャート</th>"}
                    lst.Add(vbTab & s)
                Next
            Next
            lst.Add("</tr>")

            Dim lcnt As Long = 0
            For Each k As KeyValuePair(Of String, clsTeam) In _dic
                If lcnt >= 3 Then
                    lst.Add("</tr>")
                    lcnt = 0
                End If

                If lcnt = 0 Then
                    lst.Add("<tr>")
                End If

                Dim s() As String = {"<td><center>TREAM<br><font color=""blue""><u>Facilitater</u></font></center></td>",
                    "<td><img src = ""./Image /0.png"" ></td>"}

                lst.Add("<td><center>" & k.Key & "<br><font color=""blue""><u>" & k.Value.Facilitation & "(※)</u></font></center></td>")
                lst.Add("<td><img src = ""./Image/" & k.Key & ".png" & """></td>")

                lcnt += 1
            Next
            If lcnt < 3 Then
                lst.Add("<td/>")
                lst.Add("<td/>")
            End If
            lst.Add("</tr>")
            lst.Add("</table>")


            Dim lsthtm As New List(Of String)

            Dim btable As Boolean = False
            For Each s As String In ClassLibraryUtil.clsUtil.ReadFile(tmp)
                Dim s1 As String = s.Trim.Replace(" ", "")
                If s1 = "<table>" Then
                    btable = True
                End If

                If btable = False Then
                    lsthtm.Add(s)
                Else
                    If s1 = "</table>" Then
                        lsthtm.AddRange(lst)
                        lst.Clear()
                        btable = False
                    End If
                End If
            Next

            If lsthtm.Count Then
                ClassLibraryUtil.clsUtil.WriteFile(lsthtm, htm)
            End If

            If IO.File.Exists(htm) AndAlso bShow Then
                Dim psi As New System.Diagnostics.ProcessStartInfo()
                psi.FileName = htm
                psi.WindowStyle = System.Diagnostics.ProcessWindowStyle.Maximized
                'アプリケーションを起動する
                Dim p As System.Diagnostics.Process = System.Diagnostics.Process.Start(psi)
            End If
        Loop While (False)

    End Sub

    Private Sub ChangeDgvSelect()
        Dim strTeam As String = ComboBox2.Text
        Dim strName As String = ComboBox3.Text
        With Me.DataGridView2
            Dim s As String = ""
            For Each r As DataGridViewRow In .Rows
                If String.IsNullOrEmpty(r.Cells(0).Value) = False Then
                    s = r.Cells(0).Value.ToString.Trim
                End If
                If String.IsNullOrEmpty(s) Then Continue For
                If s <> strTeam Then Continue For

                If Me.CheckBox1.Checked = False Then
                    If r.Cells(1).Value.ToString.Trim <> strName Then Continue For
                End If
                r.Selected = True
                .FirstDisplayedScrollingRowIndex = r.Index
                Exit For

            Next
        End With
    End Sub
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Herrman2Combox(DirectCast(sender, ComboBox).Text, Me.ComboBox3)
        Chart()
        ChangeDgvSelect()
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        Chart()
        ChangeDgvSelect()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        MakeTeamHtml()
    End Sub
End Class

Imports System.Data.SqlClient
Public Class Form1
    Dim sv As Object
    Dim tci As Integer
    Dim TBEC As TextBox
    Dim cmd As SqlCommand
    Dim CR, CG, CB As Byte
    Dim dr As SqlDataReader
    Dim dt As New DataTable
    Dim tb1 As New DataTable
    Dim tb2 As New DataTable
    Dim dtt As New DataTable
    Dim da As SqlDataAdapter
    Dim rc As Boolean = True
    Dim dgvcell As DataGridViewCell
    Public st() As String = Form0.st
    Public dtn, dto As New DataTable
    Public usr As String = Form0.C2.Text
    Public suer As Integer = Fcsb.s1(usr)
    Public pswd As String = Form0.T4.Text
    Dim lgc, bla, mnb, clbl, mn As Boolean
    Dim dict As New Dictionary(Of String, TextBox)
    Public lbl As New Dictionary(Of Object, Object())
    Public cmdstr, cmdstr1, cmdstr2, cmdstrgx As String
    Public dic As New Dictionary(Of String, List(Of Integer))
    Public dacw As New Dictionary(Of DataGridView, List(Of Integer))
    Public cea, ceb, ctbl, ttbl, bcbl, ccbl, nn, tcs, sbl(3), fc, BB3, BB4, L124bl As Boolean
    Public cnctk As SqlConnection = New SqlConnection("data source=" & st(3) & ";initial catalog=msdb;user id=calc")
    Public cnct As SqlConnection = New SqlConnection("data source=" & st(3) & ";initial catalog=" & st(2) & ";user id=" & usr & ";password=" & pswd)
    Public Sub B14_Click(sender As Object, e As EventArgs) Handles B14.Click
        If Not ctbl Then
            TM.Interval = SystemInformation.DoubleClickTime
            TM.Enabled = True
            ctbl = True
            Exit Sub
        Else
            TM.Interval = 1
            TM.Enabled = False
        End If
        s39(True)
    End Sub
    Private Sub B15_Click(sender As Object, e As EventArgs) Handles B15.Click
        If DA1.SelectedRows.Count = 0 Then
            DA1.Rows.Clear()
        Else
            For Each row As DataGridViewRow In DA1.SelectedRows
                If Not row.IsNewRow Then DA1.Rows.Remove(row)
            Next
        End If
    End Sub
    Private Sub B17_Click(sender As Object, e As EventArgs) Handles B17.Click
        Fcsb.s4(DA1, cnct, "学生成绩", dic)
    End Sub
    Private Sub B19_Click(sender As Object, e As EventArgs) Handles B19.Click
        lgc = True : Form0.blf = True : Close()
    End Sub
    Private Sub L126_Click(sender As Object, e As EventArgs) Handles L126.Click, B13.Click, B38.Click, L128.Click
        If Not CBool(lbl(sender)(0)) Then
            TM.Interval = SystemInformation.DoubleClickTime
            TM.Enabled = True
            lbl(sender)(0) = True
            Exit Sub
        Else
            TM.Interval = 1
            TM.Enabled = False
        End If
        s30(True, sender)
    End Sub
    Private Sub B16_Click(sender As Object, e As EventArgs) Handles B16.Click
        s7(DirectCast(sender, Button), B17, DA1, dic("学生成绩"))
    End Sub
    Private Sub B104_Click(sender As Object, e As EventArgs) Handles B104.Click
        Try
            cnct.Open()
            cmd = New SqlCommand(My.Settings.TT, cnct)
            cmd.ExecuteNonQuery()
            cnct.Close()
            MsgBox("数据整理成功！")
        Catch ex As Exception
            cnct.Close()
            MsgBox("数据整理失败！")
        End Try
    End Sub
    Private Sub B105_Click(sender As Object, e As EventArgs) Handles B105.Click
        T38.Text = "" : L5.Text = "年级："
        T52.Text = "" : T53.Text = ""
        T12.Text = "" : T15.Text = ""
        s2(LI2, LI1) : s2(LI4, LI3)
        T50.Text = "学生姓名："
        s12(LI5, L5, cmdstr1, cmdstr2)
    End Sub
    Private Sub B103_Click(sender As Object, e As EventArgs) Handles B103.Click
        Dim blct As Boolean
        If Not bcbl Then
            TM.Interval = SystemInformation.DoubleClickTime
            TM.Enabled = True
            bcbl = True
            Exit Sub
        Else
            TM.Interval = 1
            TM.Enabled = False
            blct = True
        End If
        s40(Not blct)
    End Sub
    Private Sub B35_Click(sender As Object, e As EventArgs) Handles B35.Click
        Dim bl As Boolean
        For i = 4 To DA5.Columns.Count - 1
            If CStr(DA5.Rows(0).Cells(i).Value) <> "" Then
                Try
                    cmdstr = "insert into 学生成绩 values('" & Replace(DA5.Columns(i).Name, "'", "''") & "'," & s49(CStr(DA5.Rows(0).Cells(i).Value), True, 0).Replace("'", "") & ",'" & Replace(CStr(DA5.Rows(0).Cells(3).Value), "'", "''") & "','" & Replace(CStr(DA5.Rows(0).Cells(0).Value), "'", "''") & "')select max(Id) from 学生成绩--" & DA5.Columns(i).HeaderText
                    cmd = New SqlCommand(cmdstr, cnct)
                    cnct.Open()
                    dic("学生成绩").Add(CInt(cmd.ExecuteScalar))
                    cnct.Close()
                    DA5.Rows(0).Cells(i).Value = Nothing
                    DA5.Rows(0).Cells(i).ReadOnly = True
                    bl = True
                Catch ex As Exception
                    cnct.Close()
                    MsgBox("学生:" & DA5.Columns(i).HeaderText & "的成绩输入有误。" & vbCrLf & ex.Message)
                    Exit Sub
                End Try
            End If
        Next
        If DA5.Columns.Count > 4 Then
            If bl Then
                MsgBox("数据录入成功！")
                DA5.Rows.Clear()
                For i = 4 To DA5.Columns.Count - 1
                    DA5.Columns.RemoveAt(4)
                Next
                DA5.Rows.Add()
                s56(DA5)
            Else
                MsgBox("请至少录入一项学生的成绩！")
            End If
        Else
            MsgBox("请先生成相关的学生信息！")
        End If
    End Sub
    Private Sub B34_Click(sender As Object, e As EventArgs) Handles B34.Click
        DA5.Rows.Clear()
        For i = 4 To DA5.Columns.Count - 1
            DA5.Columns.RemoveAt(4)
        Next
        DA5.Rows.Add()
        s56(DA5)
    End Sub
    Private Sub B29_Click(sender As Object, e As EventArgs) Handles B29.Click
        For Each row As DataGridViewRow In DA6.Rows
            If Not row.IsNewRow AndAlso Not row.ReadOnly Then
                Try
                    cmdstr = "insert into 学生信息 values("
                    For Each cell As DataGridViewCell In row.Cells
                        If CStr(cell.Value) = "" Then
                            cmdstr += "NULL,"
                        Else
                            cmdstr += "'" + Replace(CStr(cell.Value), "'", "''") + "',"
                        End If
                    Next
                    cmdstr += "NULL)select max(Id) from 学生信息"
                    cnct.Open()
                    cmd = New SqlCommand(cmdstr, cnct)
                    dic("学生信息").Add(CInt(cmd.ExecuteScalar()))
                    cnct.Close()
                    row.ReadOnly = True
                Catch ex As Exception
                    cnct.Close()
                    MsgBox("第" & row.Index + 1 & "行:" & ex.Message & vbCrLf & "注意：学生学号、学生姓名、入学年份、学生班级为必填项。")
                    Return
                End Try
            End If
        Next
        MsgBox("数据录入成功！")
        DA6.Rows.Clear()
    End Sub
    Private Sub B18_Click(sender As Object, e As EventArgs) Handles B18.Click
        s5(DA6)
    End Sub
    Private Sub B50_Click(sender As Object, e As EventArgs) Handles B50.Click
        DA11.Columns.Clear()
        nn = False
        cmd = New SqlCommand(T28.Text, cnct)
        Try
            cnct.Open()
            dr = cmd.ExecuteReader()
            If dr.FieldCount = 0 Then
                MsgBox("操作成功！")
            Else
                For i = 1 To dr.FieldCount
                    DA11.Columns.Add(dr.GetName(i - 1), dr.GetName(i - 1))
                Next
                While dr.Read()
                    DA11.Rows.Add()
                    For i = 1 To dr.FieldCount
                        DA11.Rows(DA11.Rows.Count - 2).Cells(i - 1).Value = IIf(IsDBNull(dr(i - 1)), Nothing, dr(i - 1))
                        DA11.Columns(i - 1).SortMode = DataGridViewColumnSortMode.Automatic
                    Next
                End While

            End If
            cnct.Close()
        Catch ex As Exception
            cnct.Close()
            MsgBox("语句查询失败！" & vbCrLf & ex.Message)
        End Try
        DA11.ClearSelection()
    End Sub
    Private Sub DA11_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DA11.CellContentClick
        Dim str1, ph As String : Dim bl As Boolean
        Dim str2 As Integer
        Dim DA As DataGridView = DirectCast(sender, DataGridView)
        If e.RowIndex < 0 OrElse DA.Columns.Count < 4 Then Exit Sub
        str1 = CStr(DA.Rows(e.RowIndex).Cells(3).Value)
        If nn Then
            If e.ColumnIndex = 4 Then
                str2 = CInt(DA.Rows(e.RowIndex).Cells(4).Value)
                Select Case str1
                    Case "学生成绩"
                        s8(str2)
                        DA1.ClearSelection()
                    Case "任课信息"
                        s9(str2)
                        DGV3.ClearSelection()
                    Case "学生信息"
                        s13("Id", CObj(str2))
                End Select
            End If
        End If
    End Sub
    Private Sub DA11_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DA11.CellMouseClick
        Dim DA As DataGridView = DirectCast(sender, DataGridView)
        If e.RowIndex = -1 Then
            If e.Button = Windows.Forms.MouseButtons.Middle Then
                For i = 1 To DA.Columns.Count
                    DA.Columns.Item(i - 1).Visible = True
                Next
            ElseIf e.Button = Windows.Forms.MouseButtons.Right AndAlso e.ColumnIndex > -1 Then
                DA.Columns.Item(e.ColumnIndex).Visible = False
            End If
        ElseIf e.ColumnIndex > -1 Then
            If nn Then
                If e.ColumnIndex = 4 Then
                    If e.Button = Windows.Forms.MouseButtons.Right Then
                        s25(DA, CStr(DA.Rows(e.RowIndex).Cells(4).Value), CStr(DA.Rows(e.RowIndex).Cells(3).Value))
                    End If
                End If
            End If
        End If
    End Sub
    Private Sub B51_Click(sender As Object, e As EventArgs) Handles B51.Click
        s8(DA11)
    End Sub
    Private Sub B52_Click(sender As Object, e As EventArgs) Handles B52.Click
        If sbl(3) Then
        ElseIf Not ccbl Then
            TM.Interval = SystemInformation.DoubleClickTime
            TM.Enabled = True
            ccbl = True
            Exit Sub
        Else
            TM.Interval = 1
            TM.Enabled = False
        End If
        s4(True)
        s56(DA11)
    End Sub
    Public Sub DA1_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles DA1.CellBeginEdit
        s36(e, DirectCast(sender, DataGridView), "select 学生成绩.*,学生信息.学生姓名 from 学生成绩 full outer join 学生信息 on 学生成绩.学生学号=学生信息.学生学号 where 学生成绩.Id=", sv)
    End Sub
    Public Sub DA1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DA1.CellEndEdit
        Dim num As Decimal
        Dim flag As Boolean
        Dim DA As DataGridView = DirectCast(sender, DataGridView)
        rc = True
        cmdstr = ""
        If DA.Rows.Count = 1 OrElse CStr(DA.Rows(DA.Rows.Count - 2).Cells(0).Value) <> "" Then
            For i As Integer = 1 To DA.Columns.Count
                DA.Columns(i - 1).SortMode = DataGridViewColumnSortMode.Automatic
            Next
        End If
        If CStr(DA.Rows(e.RowIndex).Cells(0).Value) = "" Then
            If e.ColumnIndex = 1 Then
                Try
                    cnct.Open()
                    DA.Rows(e.RowIndex).Cells(5).Value = New SqlCommand("select 学生姓名 from 学生信息 where 学生学号='" & CStr(DA.Rows(e.RowIndex).Cells(1).Value).Replace("'", "''") & "'", cnct).ExecuteScalar
                    cnct.Close()
                Catch ex As Exception
                    cnct.Close()
                End Try
            End If
        ElseIf CInt(DA.Rows(e.RowIndex).Cells(0).Value) > 0 Then
            CR = DA.Rows(e.RowIndex).Cells(0).Style.BackColor.R
            CG = DA.Rows(e.RowIndex).Cells(0).Style.BackColor.G
            CB = DA.Rows(e.RowIndex).Cells(0).Style.BackColor.B
            mn = False
            cea = False
            If e.ColumnIndex = 1 Then
                If Not IsNumeric(CStr(DA.Rows(e.RowIndex).Cells(1).Value)) OrElse Len(CStr(DA.Rows(e.RowIndex).Cells(1).Value)) <> 8 Then
                    MsgBox("学生学号输入有误，请检查后重输！", MsgBoxStyle.OkOnly, Nothing)
                    s11(e)
                    tcs = True
                    rc = False
                    Return
                End If
                mn = False
                cea = False
                tcs = False
                If CStr(sv) <> CStr(DA.Rows(e.RowIndex).Cells(1).Value) Then cmdstr = String.Concat("update 学生成绩 set 学生学号='", CStr(DA.Rows(e.RowIndex).Cells(1).Value), "' where Id=", CStr(DA.Rows(e.RowIndex).Cells(0).Value))
            ElseIf e.ColumnIndex = 2 Then
                DA.Rows(e.RowIndex).Cells(2).Value = s49(CStr(DA.Rows(e.RowIndex).Cells(2).Value), flag, num)
                If Not flag Then
                    MsgBox("学生成绩输入有误，请检查后重输！", MsgBoxStyle.OkOnly, Nothing)
                    s11(e)
                    tcs = True
                    rc = False
                    Return
                End If
                mn = False
                cea = False
                tcs = False
                If CDec(Format(num, "0.0")) <> CDec(sv) OrElse Not IsNumeric(DA.Rows(e.RowIndex).Cells(2).Value) Then
                    cmdstr = String.Concat(New String() {"update 学生成绩 set 学生成绩=", CStr(DA.Rows(e.RowIndex).Cells(2).Value), " where Id=", CStr(DA.Rows(e.RowIndex).Cells(0).Value)})
                End If
                DA.Rows(e.RowIndex).Cells(2).Value = CDec(Format(num, "0.0"))
            ElseIf CStr(DA.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) <> CStr(sv) Then
                If CStr(DA.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) <> "" Then
                    cmdstr = String.Concat(New String() {"update 学生成绩 set ", DA.Columns(e.ColumnIndex).HeaderText, "='", Replace(CStr(DA.Rows(e.RowIndex).Cells(e.ColumnIndex).Value), "'", "''"), "' where Id=", CStr(DA.Rows(e.RowIndex).Cells(0).Value)})
                Else
                    MsgBox(DA.Columns(e.ColumnIndex).HeaderText & "不能为空！")
                    s11(e)
                    tcs = True
                    rc = False
                    Return
                End If
            End If
            If cmdstr <> "" Then
                Try
                    cmd = New SqlCommand(cmdstr, cnct)
                    cnct.Open()
                    cmd.ExecuteNonQuery()
                    If CInt(DA.Rows(e.RowIndex).Cells(0).Value) <> 0 Then
                        If CR = 255 AndAlso CG = 255 AndAlso CB = 255 Then
                            DA.Rows(e.RowIndex).Cells(0).Style.BackColor = Color.Red
                        ElseIf CR = 255 Then
                            DA.Rows(e.RowIndex).Cells(0).Style.BackColor = Color.FromArgb(0, 255, 0)
                        ElseIf CG <> 255 Then
                            DA.Rows(e.RowIndex).Cells(0).Style.BackColor = Color.Red
                            DA.Rows(e.RowIndex).Cells(0).Style.ForeColor = Color.Black
                        Else
                            DA.Rows(e.RowIndex).Cells(0).Style.BackColor = Color.Blue
                            DA.Rows(e.RowIndex).Cells(0).Style.ForeColor = Color.FromArgb(255, 255, 0)
                        End If
                    End If
                    cnct.Close()
                    If Not dic("学生成绩").Contains(CInt(DA.Rows(e.RowIndex).Cells(0).Value)) Then
                        dic("学生成绩").Add(CInt(DA.Rows(e.RowIndex).Cells(0).Value))
                    End If
                    mn = False
                    cea = False
                    tcs = False
                Catch ex As Exception
                    cnct.Close()
                    MsgBox(String.Concat("数据更新出错了..." & vbCrLf & "", ex.Message), MsgBoxStyle.OkOnly, Nothing)
                    s11(e)
                    tcs = True
                    rc = False
                    Return
                End Try
                rc = True
            End If
        End If
    End Sub
    Public Sub DA1_RowValidating(sender As Object, e As DataGridViewCellCancelEventArgs) Handles DA1.RowValidating
        Dim num As Decimal
        Dim count As Integer
        Dim DA As DataGridView = DirectCast(sender, DataGridView)
        If DA.NewRowIndex <> e.RowIndex AndAlso Not DA.Rows(e.RowIndex).ReadOnly AndAlso CStr(DA.Rows(e.RowIndex).Cells(0).Value) = "" Then
            DA.EndEdit()
            Dim str(2) As String
            Dim flag As Boolean = True
            str(0) = Replace(CStr(DA.Rows(e.RowIndex).Cells(1).Value), "'", "''")
            str(1) = Replace(CStr(DA.Rows(e.RowIndex).Cells(3).Value), "'", "''")
            str(2) = Replace(CStr(DA.Rows(e.RowIndex).Cells(4).Value), "'", "''")
            DA.Rows(e.RowIndex).Cells(2).Value = Replace(CStr(DA.Rows(e.RowIndex).Cells(2).Value), "'", "")
            DA.Rows(e.RowIndex).Cells(2).Value = s49(CStr(DA.Rows(e.RowIndex).Cells(2).Value), flag, num)
            Dim str1 As String = "insert into 学生成绩 values("
            Dim str2 As String = "'" + str(0) + "'"
            Dim str3 As String = CStr(DA.Rows(e.RowIndex).Cells(2).Value)
            Dim str4 As String = "'" + str(1) + "'"
            Dim str5 As String = "'" + str(2) + "'"
            If CStr(DA.Rows(e.RowIndex).Cells(1).Value) = "" Then
                s16(DA, 1, e)
                Return
            End If
            If Not flag OrElse CStr(DA.Rows(e.RowIndex).Cells(2).Value) = "" Then
                s16(DA, 2, e)
                Return
            End If
            If CStr(DA.Rows(e.RowIndex).Cells(3).Value) = "" Then
                s16(DA, 3, e)
                Return
            End If
            If CStr(DA.Rows(e.RowIndex).Cells(4).Value) = "" Then
                s16(DA, 4, e)
                Return
            End If
            cmdstr = String.Concat(str1, str2, ",", str3, ",", str4, ",", str5, ")")
            count = DA.Rows.Count - 2
            If Not Fcsb.s9(count, DA, "学生成绩", cnct, cmdstr) Then
                dic("学生成绩").Add(CInt(DA1.Rows(e.RowIndex).Cells(0).Value))
                DA.Rows(e.RowIndex).Cells(0).Style.BackColor = Color.DarkViolet
                DA.Rows(e.RowIndex).Cells(0).Style.ForeColor = Color.White
            End If
            tcs = False
            count = DA.Columns.Count
            Dim num2 As Integer = 1
            Do
                DA.Columns(num2 - 1).SortMode = DataGridViewColumnSortMode.Automatic
                num2 = num2 + 1
            Loop While num2 <= count
            DA1.Rows(e.RowIndex).Cells(2).Value = CDec(Format(num, "0.0"))
        End If
    End Sub
    Public Sub DGV3_RowValidating(sender As Object, e As DataGridViewCellCancelEventArgs) Handles DGV3.RowValidating
        Dim count As Integer
        Dim DA As DataGridView = DirectCast(sender, DataGridView)
        If DA.NewRowIndex <> e.RowIndex AndAlso Not DA.Rows(e.RowIndex).ReadOnly AndAlso CStr(DA.Rows(e.RowIndex).Cells(0).Value) = "" Then
            DA.EndEdit()
            Dim str(3) As String
            str(0) = Replace(CStr(DA.Rows(e.RowIndex).Cells(1).Value), "'", "''")
            str(1) = Replace(CStr(DA.Rows(e.RowIndex).Cells(2).Value), "'", "''")
            str(2) = Replace(CStr(DA.Rows(e.RowIndex).Cells(3).Value), "'", "''")
            str(3) = Replace(CStr(DA.Rows(e.RowIndex).Cells(4).Value), "'", "''")
            Dim str1 As String = "insert into 任课信息 values("
            Dim str2 As String = "'" + str(0) + "'"
            Dim str3 As String = "'" + str(1) + "'"
            Dim str4 As String = "'" + str(2) + "'"
            Dim str5 As String = "'" + str(3) + "'"
            If CStr(DA.Rows(e.RowIndex).Cells(1).Value) = "" Then
                s16(DA, 1, e)
                Return
            ElseIf Not IsNumeric(DA.Rows(e.RowIndex).Cells(1).Value) OrElse Len(CStr(DA.Rows(e.RowIndex).Cells(1).Value)) <> 4 Then
                s16(DA, 1, e)
                Return
            End If
            If CStr(DA.Rows(e.RowIndex).Cells(2).Value) = "" Then
                s16(DA, 2, e)
                Return
            ElseIf Not IsNumeric(DA.Rows(e.RowIndex).Cells(2).Value) Then
                s16(DA, 2, e)
                Return
            End If
            If CStr(DA.Rows(e.RowIndex).Cells(3).Value) = "" Then
                s16(DA, 3, e)
                Return
            End If
            If CStr(DA.Rows(e.RowIndex).Cells(4).Value) = "" Then
                s16(DA, 4, e)
                Return
            End If
            cmdstr = String.Concat(str1, str2, ",", str3, ",", str4, ",", str5, ")")
            count = DA.Rows.Count - 2
            If Not Fcsb.s9(count, DA, "任课信息", cnct, cmdstr) Then
                dic("任课信息").Add(CInt(DGV3.Rows(e.RowIndex).Cells(0).Value))
                DA.Rows(e.RowIndex).Cells(0).Style.BackColor = Color.Pink
                DA.Rows(e.RowIndex).Cells(0).Style.ForeColor = Color.White
                DA.Rows(e.RowIndex).Cells(1).Value = CShort(DA.Rows(e.RowIndex).Cells(1).Value)
                DA.Rows(e.RowIndex).Cells(2).Value = CByte(DA.Rows(e.RowIndex).Cells(2).Value)
            End If
            tcs = False
            count = DA.Columns.Count
            Dim num2 As Integer = 1
            Do
                DA.Columns(num2 - 1).SortMode = DataGridViewColumnSortMode.Automatic
                num2 = num2 + 1
            Loop While num2 <= count
        End If
    End Sub
    Private Sub DA1_KeyDown(sender As Object, e As KeyEventArgs) Handles DA1.KeyDown
        Dim DA As DataGridView = DirectCast(sender, DataGridView)
        If e.KeyCode = Keys.Escape Then
            Try
                If DA.Rows(DA.Rows.Count - 2).Cells(0).Value Is Nothing Then
                    RemoveHandler DA.RowValidating, AddressOf DA1_RowValidating
                    DA.Rows.RemoveAt(DA.Rows.Count - 2)
                    tcs = False
                    AddHandler DA.RowValidating, AddressOf DA1_RowValidating
                End If
                rc = True
            Catch ex As Exception
            End Try
        End If
    End Sub
    Private Sub DGV3_KeyDown(sender As Object, e As KeyEventArgs) Handles DGV3.KeyDown
        Dim DA As DataGridView = DirectCast(sender, DataGridView)
        If e.KeyCode = Keys.Escape Then
            Try
                If DA.Rows(DA.Rows.Count - 2).Cells(0).Value Is Nothing Then
                    RemoveHandler DA.RowValidating, AddressOf DGV3_RowValidating
                    DA.Rows.RemoveAt(DA.Rows.Count - 2)
                    tcs = False
                    AddHandler DA.RowValidating, AddressOf DGV3_RowValidating
                End If
                rc = True
            Catch ex As Exception
            End Try
        End If
    End Sub
    Private Sub RowPostPaint(sender As Object, e As DataGridViewRowPostPaintEventArgs) Handles DA1.RowPostPaint, DGV3.RowPostPaint, DGV4.RowPostPaint, DA2.RowPostPaint, DA5.RowPostPaint, DA6.RowPostPaint, DA11.RowPostPaint
        If suer <> 4 Then TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(), New Font("Times New Roman", 9), New Rectangle(e.RowBounds.Location.X, e.RowBounds.Location.Y + 4, DirectCast(sender, DataGridView).RowHeadersWidth - 4, e.RowBounds.Height), DirectCast(sender, DataGridView).RowHeadersDefaultCellStyle.ForeColor, Color.Transparent, TextFormatFlags.HorizontalCenter)
    End Sub
    Sub s11(ByRef e As DataGridViewCellEventArgs)
        RemoveHandler DA1.RowValidating, AddressOf DA1_RowValidating
        RemoveHandler DA1.CellMouseClick, AddressOf DA1_CellMouseClick
        RemoveHandler DA1.CellMouseUp, AddressOf DA1_CellMouseUp
        RemoveHandler DA1.CellEndEdit, AddressOf DA1_CellEndEdit
        RemoveHandler DA1.CellBeginEdit, AddressOf DA1_CellBeginEdit
        DA1.Columns(e.ColumnIndex).Visible = True
        DA1.CurrentCell = DA1.Rows(e.RowIndex).Cells(e.ColumnIndex)
        DA1.BeginEdit(False) : cea = True
        dgvcell = DA1.Rows(e.RowIndex).Cells(e.ColumnIndex) : mn = True
        For i = 1 To DA1.Columns.Count
            DA1.Columns(i - 1).SortMode = DataGridViewColumnSortMode.NotSortable
        Next
        AddHandler DA1.CellBeginEdit, AddressOf DA1_CellBeginEdit
        AddHandler DA1.CellEndEdit, AddressOf DA1_CellEndEdit
        AddHandler DA1.CellMouseUp, AddressOf DA1_CellMouseUp
        AddHandler DA1.CellMouseClick, AddressOf DA1_CellMouseClick
        AddHandler DA1.RowValidating, AddressOf DA1_RowValidating
    End Sub
    Sub s36(ByRef e As DataGridViewCellCancelEventArgs, ByRef DA As DataGridView, ByRef cmdstr As String, ByRef sv As Object)
        rc = False : sv = Nothing
        For i = 1 To DA.Columns.Count
            DA.Columns(i - 1).SortMode = DataGridViewColumnSortMode.NotSortable
        Next
        If CStr(DA.Rows(e.RowIndex).Cells(0).Value) = "" Then Return
        cmdstr += CStr(DA.Rows(e.RowIndex).Cells(0).Value)
        cmd = New SqlCommand(cmdstr, cnct)
        Try
            cnct.Open()
            dr = cmd.ExecuteReader
            If dr.HasRows Then
                While dr.Read
                    sv = IIf(IsDBNull(dr(e.ColumnIndex)), Nothing, dr(e.ColumnIndex))
                    For Each col As DataGridViewColumn In DA.Columns
                        If TypeOf col Is DataGridViewComboBoxColumn AndAlso Not DirectCast(col, DataGridViewComboBoxColumn).Items.Contains(CStr(IIf(IsDBNull(dr(col.Index)), "", dr(col.Index)))) Then
                            MsgBox("数据库值:" & CStr(IIf(IsDBNull(dr(col.Index)), Nothing, dr(col.Index))) & " 不在" & col.HeaderText & "列表中！")
                        Else
                            DA.Rows(e.RowIndex).Cells(col.Index).Value = IIf(IsDBNull(dr(col.Index)), Nothing, dr(col.Index))
                        End If
                    Next
                End While
            Else
                sv = Nothing
                e.Cancel = True
                Dim RowData(1) As Integer
                RowData(0) = Nothing
                RowData(1) = CInt(DA.Rows(e.RowIndex).Cells(0).Value)
                DA.Rows(e.RowIndex).Cells(0).Value = 0
                DA.Rows(e.RowIndex).ReadOnly = True
                DA.Rows(e.RowIndex).Tag = RowData
            End If
            cnct.Close()
        Catch ex As Exception
            cnct.Close()
        End Try
    End Sub
    Private Sub B20_Click(sender As Object, e As EventArgs) Handles B20.Click
        Try
            If B31.Text <> "" OrElse T2.Text <> "" Then
                cnct.Open()
                If B31.Text <> "" Then
                    cmd = New SqlCommand("select * from 学生信息 where Id=" & B31.Text, cnct)
                    Dim dt As New DataTable
                    Dim da As SqlDataAdapter = New SqlDataAdapter(cmd)
                    da.Fill(dt)
                    For Each key As String In dict.Keys
                        If Trim(dt.Rows(0)(dt.Columns.Item(key)).ToString) <> Trim(dict(key).Text) Then
                            cmd = New SqlCommand("update 学生信息 set " & key & "=" & CStr(IIf(Trim(dict(key).Text) = "", "NULL", "'" & Replace(Trim(dict(key).Text), "'", "''") & "'")) & " where Id=" & B31.Text, cnct)
                            cmd.ExecuteNonQuery()
                        End If
                    Next
                ElseIf T2.Text <> "" Then
                    cmdstr = "insert into 学生信息 values("
                    For Each key As String In dict.Keys
                        cmdstr += CStr(IIf(Trim(dict(key).Text) = "", "NULL", "'" & Replace(Trim(dict(key).Text), "'", "''") & "'")) & ","
                    Next
                    cmdstr = Strings.Left(cmdstr, Len(cmdstr) - 1) & ")select max(Id) from 学生信息"
                    cmd = New SqlCommand(cmdstr, cnct)
                    B31.Text = CStr(cmd.ExecuteScalar)
                    dic("学生信息").Add(CInt(B31.Text))
                End If
                cnct.Close()
                T2.Tag = T2.Text
                MsgBox("学生信息保存成功！")
            End If
        Catch ex As Exception
            cnct.Close()
            MsgBox("学生信息保存失败！,具体查看操作记录" & vbCrLf & ex.Message & vbCrLf & "注意：学生学号、学生姓名、入学年份、学生班级为必填项。")
        End Try
    End Sub
    Private Sub B21_Click(sender As Object, e As EventArgs) Handles B21.Click
        If B31.Text <> "" Then
            cmd = New SqlCommand("delete from 学生信息 where Id=" & B31.Text, cnct)
            Try
                cnct.Open()
                cmd.ExecuteNonQuery()
                dic("学生信息").Remove(CInt(B31.Text))
                B31.Text = ""
                cnct.Close()
                MsgBox("学生信息删除成功！")
            Catch ex As Exception
                cnct.Close()
                MsgBox("学生信息删除失败！")
            End Try
        End If
    End Sub
    Private Sub B1_Click(sender As Object, e As EventArgs) Handles B1.Click, LI1.DoubleClick
        s1(LI1, LI2)
    End Sub
    Private Sub B2_Click(sender As Object, e As EventArgs) Handles B2.Click, LI2.DoubleClick
        s1(LI2, LI1)
    End Sub
    Private Sub B5_Click(sender As Object, e As EventArgs) Handles B5.Click, LI3.DoubleClick
        s1(LI3, LI4)
    End Sub
    Private Sub B6_Click(sender As Object, e As EventArgs) Handles B6.Click, LI4.DoubleClick
        s1(LI4, LI3)
    End Sub
    Private Sub B3_Click(sender As Object, e As EventArgs) Handles B3.Click
        s2(LI1, LI2)
    End Sub
    Private Sub B4_Click(sender As Object, e As EventArgs) Handles B4.Click
        s2(LI2, LI1)
    End Sub
    Private Sub B7_Click(sender As Object, e As EventArgs) Handles B7.Click
        s2(LI3, LI4)
    End Sub
    Private Sub B8_Click(sender As Object, e As EventArgs) Handles B8.Click
        s2(LI4, LI3)
    End Sub
    Sub s1(ByRef li1 As ListBox, ByRef li2 As ListBox)
        If IsNothing(li1.SelectedItem) Then Exit Sub
        For Each r In li1.SelectedItems
            If Not li2.Items.Contains(r) Then li2.Items.Add(r)
        Next
        For i = 0 To li1.SelectedItems.Count - 1
            li1.Items.Remove(li1.SelectedItems.Item(0))
        Next
    End Sub
    Sub s2(ByRef li1 As ListBox, ByRef li2 As ListBox)
        If li1.Items.Count = 0 Then Exit Sub
        For i = 0 To li1.Items.Count - 1
            If Not li2.Items.Contains(li1.Items.Item(0)) Then li2.Items.Add(li1.Items.Item(0))
            li1.Items.RemoveAt(0)
        Next i
    End Sub
    Private Sub L_LostFocus(sender As Object, e As EventArgs) Handles L8.LostFocus, L10.LostFocus, L12.LostFocus, L27.LostFocus, L48.LostFocus, T15.LostFocus
        AcceptButton = B14
    End Sub
    Private Sub LT_Keyup(sender As Object, e As KeyEventArgs) Handles L8.KeyUp, L10.KeyUp, L12.KeyUp, L27.KeyUp, L48.KeyUp, L16.KeyUp, L18.KeyUp, L20.KeyUp, T54.KeyUp, T55.KeyUp, T56.KeyUp, T57.KeyUp, T58.KeyUp, T52.KeyUp, T53.KeyUp
        RemoveHandler DirectCast(sender, TextBox).TextChanged, AddressOf LT_TextChanged
        s51(DirectCast(sender, TextBox), e, cnctk, True)
        AddHandler DirectCast(sender, TextBox).TextChanged, AddressOf LT_TextChanged
    End Sub
    Public Sub LT_TextChanged(sender As Object, e As EventArgs) Handles L8.TextChanged, L10.TextChanged, L12.TextChanged, L27.TextChanged, L48.TextChanged, L16.TextChanged, L18.TextChanged, L20.TextChanged, T54.TextChanged, T55.TextChanged, T56.TextChanged, T57.TextChanged, T58.TextChanged, T52.TextChanged, T53.TextChanged
        DirectCast(sender, Control).Tag = DirectCast(sender, Control).Text
    End Sub
    Private Sub LT_GotFocus(sender As Object, e As EventArgs) Handles L8.GotFocus, L10.GotFocus, L12.GotFocus, L27.GotFocus, L48.GotFocus, L16.GotFocus, L18.GotFocus, L20.GotFocus, T54.GotFocus, T55.GotFocus, T56.GotFocus, T57.GotFocus, T58.GotFocus, T52.GotFocus, T53.GotFocus
        AcceptButton = Nothing
    End Sub
    Private Sub TM_Tick(sender As Object, e As EventArgs) Handles TM.Tick
        DirectCast(sender, Timer).Enabled = False
        DirectCast(sender, Timer).Interval = 1
        If ctbl Then
            RemoveHandler B14.Click, AddressOf B14_Click
            s39(CBool(IIf(suer = 4, True, False)))
            AddHandler B14.Click, AddressOf B14_Click
        ElseIf BB3 Then
            RemoveHandler Form2.B3.Click, AddressOf Form2.B3_Click
            s30(False, DirectCast(Form2.B3, Object), False)
            AddHandler Form2.B3.Click, AddressOf Form2.B3_Click
        ElseIf BB4 Then
            RemoveHandler Form2.B6.Click, AddressOf Form2.B6_Click
            s30(False, DirectCast(Form2.B6, Object), False)
            AddHandler Form2.B6.Click, AddressOf Form2.B6_Click
        ElseIf bcbl Then
            RemoveHandler B103.Click, AddressOf B103_Click
            s40(True)
            AddHandler B103.Click, AddressOf B103_Click
        ElseIf ccbl Then
            RemoveHandler B52.Click, AddressOf B52_Click
            s4(False)
            AddHandler B52.Click, AddressOf B52_Click
            s56(DA11)
        ElseIf ttbl Then
            RemoveHandler B11.Click, AddressOf B11_Click
            s38(False)
            AddHandler B11.Click, AddressOf B11_Click
        ElseIf CBool(lbl(L126)(0)) Then
            RemoveHandler L126.Click, AddressOf L126_Click
            s30(False, DirectCast(L126, Object))
            AddHandler L126.Click, AddressOf L126_Click
        ElseIf CBool(lbl(B13)(0)) Then
            RemoveHandler B13.Click, AddressOf L126_Click
            s30(False, DirectCast(B13, Object))
            AddHandler B13.Click, AddressOf L126_Click
        ElseIf CBool(lbl(B38)(0)) Then
            RemoveHandler B38.Click, AddressOf L126_Click
            s30(False, DirectCast(B38, Object))
            AddHandler B38.Click, AddressOf L126_Click
        ElseIf CBool(lbl(L128)(0)) Then
            RemoveHandler L128.Click, AddressOf L126_Click
            s30(False, DirectCast(L128, Object))
            AddHandler L128.Click, AddressOf L126_Click
        ElseIf bla Then
            Dim str As Short
            If Len(CStr(DA5.Rows(0).Cells(2).Value)) = 4 Then
                str = CShort(DA5.Rows(0).Cells(2).Value)
            ElseIf Len(CStr(DA5.Rows(0).Cells(2).Value)) = 1 Then
                str = Fcsb.s3(DA5.Rows(0).Cells(2).Value, DBNull.Value)
            Else
                AddHandler DA5.CellEndEdit, AddressOf DA5_CellEndEdit
                bla = False
                Return
            End If
            For i = 4 To DA5.Columns.Count - 1
                DA5.Columns.RemoveAt(4)
            Next
            cmdstr = "select 学生学号,学生姓名 from 学生信息 where 入学年份=@入学年份 and 学生班级=@学生班级"
            Try
                cnct.Open()
                cmd = New SqlCommand(cmdstr, cnct)
                cmd.Parameters.AddWithValue("@入学年份", str)
                cmd.Parameters.AddWithValue("@学生班级", DA5.Rows(0).Cells(1).Value)
                dr = cmd.ExecuteReader
                While dr.Read
                    DA5.Columns.Add(CStr(dr(0)), CStr(dr(1)))
                End While
                dr.Close()
            Catch ex As Exception
                dr.Close()
            End Try
            cnct.Close()
            AddHandler DA5.CellEndEdit, AddressOf DA5_CellEndEdit
            bla = False
            s56(DA5)
        ElseIf L124bl Then
            RemoveHandler Form2.B7.MouseClick, AddressOf Form2.B7_MouseClick
            Form2.s22(False, Form2.ex)
            AddHandler Form2.B7.MouseClick, AddressOf Form2.B7_MouseClick
        End If
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Text += "-" & usr
        If Form0.CL1.CheckedItems.Count = 0 Then clbl = True
        sbl(0) = suer = 0 OrElse suer = 1
        sbl(1) = suer = 2 OrElse suer = 4
        sbl(2) = suer = 4
        sbl(3) = suer = 3
        DA5.Rows.Add()
        For o = 0 To 3
            DA5.Columns(o).Frozen = True
        Next
        If clbl Then
            For Each r As String In Form0.CL1.Items
                LI3.Items.Add(r)
                LI7.Items.Add(r)
                DirectCast(DA1.Columns(3), DataGridViewComboBoxColumn).Items.Add(r)
                DirectCast(DA5.Columns(3), DataGridViewComboBoxColumn).Items.Add(r)
                DirectCast(DGV3.Columns(3), DataGridViewComboBoxColumn).Items.Add(r)
                DA2.Columns.Add(r, r)
                DA2.Columns(r).Width = 70
            Next
        Else
            For Each r As String In Form0.CL1.CheckedItems
                LI3.Items.Add(r)
                LI7.Items.Add(r)
                DirectCast(DA1.Columns(3), DataGridViewComboBoxColumn).Items.Add(r)
                DirectCast(DA5.Columns(3), DataGridViewComboBoxColumn).Items.Add(r)
                DirectCast(DGV3.Columns(3), DataGridViewComboBoxColumn).Items.Add(r)
                DA2.Columns.Add(r, r)
                DA2.Columns(r).Width = 70
            Next
        End If
        DA2.Columns.RemoveAt(2)
        LI3.Items.Remove("全部")
        LI7.Items.Remove("全部")
        DirectCast(DA1.Columns(3), DataGridViewComboBoxColumn).Items.Remove("全部")
        DirectCast(DA5.Columns(3), DataGridViewComboBoxColumn).Items.Remove("全部")
        DirectCast(DGV3.Columns(3), DataGridViewComboBoxColumn).Items.Remove("全部")
        DirectCast(DA5.Columns(3), DataGridViewComboBoxColumn).Items.Remove("全部")
        cmdstr1 = "select dbo.年份换算(NULL,7,NULL)入学年份,7学生年级 union select dbo.年份换算(NULL,8,NULL)入学年份,8学生年级 union select dbo.年份换算(NULL,9,NULL)入学年份,9学生年级 order by 学生年级"
        cmdstr2 = "select 学生姓名,学生学号 from 学生信息 where 入学年份 in(dbo.年份换算(NULL,7,NULL),dbo.年份换算(NULL,8,NULL),dbo.年份换算(NULL,9,NULL))order by 学生学号"
        s12(LI5, L5, cmdstr1, cmdstr2)
        s12(LI9, T11, cmdstr1, "")
        cmdstr = "select distinct 任课教师 from 任课信息 where 考试科目 in("
        For Each item In LI3.Items
            cmdstr += "'" & item.ToString.Replace("'", "''") & "',"
        Next
        cmdstr = Strings.Left(cmdstr, Len(cmdstr) - 1) & ")"
        cmd = New SqlCommand(cmdstr, cnct)
        Try
            cnct.Open()
            dr = cmd.ExecuteReader
            Dim i As Integer
            dtt.Columns.Add("任课教师")
            dtt.Columns.Add("Id", Type.GetType("System.Int32"))
            While dr.Read
                LI10.Items.Add(dr(0))
                dtt.Rows.Add(dr(0), i)
                i += 1
            End While
            cnct.Close()
        Catch ex As Exception
            cnct.Close()
        End Try
        If suer = 4 Then
            L128.Enabled = False
            L126.Enabled = False
            B13.Enabled = False
            B38.Enabled = False
        End If
        lbl.Add(L126, {False, DA1})
        lbl.Add(L128, {False, DA11})
        lbl.Add(B13, {False, DGV3})
        lbl.Add(B38, {False, DGV4})
        For Each dtr As DataRow In dtn.Rows
            If Not LI1.Items.Contains(dtr(0)) Then LI1.Items.Add(dtr(0))
        Next
        D9.Text = Format(DateAdd(DateInterval.Day, -1, Now), "yyyy-MM-dd 00:00")
        D10.Text = Format(DateAdd(DateInterval.Day, 1, Now), "yyyy-MM-dd 00:00")
        If suer = 1 Then
            B104.Enabled = False
            B103.Enabled = False
        End If
        If suer = 2 Then
            B104.Enabled = False
            B103.Enabled = False
            B50.Enabled = False
        End If
        If suer = 3 Then
            B22.Enabled = False
            B23.Enabled = False
            B16.Enabled = False
            B17.Enabled = False
            B104.Enabled = False
            B103.Enabled = False
            B9.Enabled = False
            B29.Enabled = False
            B35.Enabled = False
            B50.Enabled = False
        End If
        If suer = 4 Then
            L126.Enabled = False
            B104.Enabled = False
            B103.Enabled = False
            B13.Enabled = False
            B38.Enabled = False
            B50.Enabled = False
            L128.Enabled = False
            DA11.ReadOnly = True
            DA1.ClipboardCopyMode = DataGridViewClipboardCopyMode.Disable
            DGV3.ClipboardCopyMode = DataGridViewClipboardCopyMode.Disable
            DGV4.ClipboardCopyMode = DataGridViewClipboardCopyMode.Disable
            DA11.ClipboardCopyMode = DataGridViewClipboardCopyMode.Disable
        End If
        dacw.Add(DA1, New List(Of Integer)) : s56(DA1)
        dacw.Add(DA2, New List(Of Integer)) : s56(DA2)
        dacw.Add(DA5, New List(Of Integer)) : s56(DA5)
        dacw.Add(DA6, New List(Of Integer)) : s56(DA6)
        dacw.Add(DGV3, New List(Of Integer)) : s56(DGV3)
        dacw.Add(DGV4, New List(Of Integer)) : s56(DGV4)
        dacw.Add(DA11, New List(Of Integer))
        Dim id As List(Of Integer)
        id = New List(Of Integer)
        dic.Add("学生成绩", id)
        id = New List(Of Integer)
        dic.Add("学生信息", id)
        id = New List(Of Integer)
        dic.Add("任课信息", id)
        dict.Add("学生学号", T2) : dict.Add("学生姓名", T1)
        dict.Add("入学年份", T9) : dict.Add("学生班级", T8)
        dict.Add("学生性别", T3) : dict.Add("学生父亲", T4)
        dict.Add("父亲电话", T5) : dict.Add("学生母亲", T6)
        dict.Add("母亲电话", T7) : dict.Add("学生备注", T10)
        If sbl(1) Then
            T38.Enabled = False
        Else
            T38.Text = Format(DateAdd(DateInterval.Month, -2, Now()), "yyyyMM") & vbCrLf & Format(DateAdd(DateInterval.Month, -1, Now()), "yyyyMM") & vbCrLf & Format(Now(), "yyyyMM")
        End If
    End Sub
    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Dim na, nb, nc As Boolean
        na = DA2.Rows.Count > 1
        nb = DA6.Rows.Count > 1
        nc = DA5.Columns.Count > 4
        If Not nc Then
            For Each cell As DataGridViewCell In DA6.Rows(0).Cells
                If CStr(cell.Value) <> "" Then
                    nc = True
                    Exit For
                End If
            Next
        End If
        If (na OrElse nb OrElse nc) AndAlso Not fc Then
            TC1.SelectedIndex = 2
            Dim msgr As MsgBoxResult = MsgBox("有未提交的行，是否继续退出", MsgBoxStyle.OkCancel)
            If msgr = MsgBoxResult.Cancel Then
                e.Cancel = True : lgc = False : Exit Sub
            ElseIf lgc Then
                Hide()
                Form0.Show()
            End If
        ElseIf lgc Then
            lgc = False
            Hide()
            Form0.Show()
        End If
        Form2.Close()
    End Sub
    Private Sub B9_Click(sender As Object, e As EventArgs) Handles B9.Click
        Dim i As Short
        Dim j As Byte
        Dim bl As Boolean
        For Each row As DataGridViewRow In DA2.Rows
            If Not row.IsNewRow Then
                If Short.TryParse(CStr(row.Cells(0).Value), i) AndAlso Byte.TryParse(CStr(row.Cells(1).Value), j) Then
                    For Each cell As DataGridViewCell In row.Cells
                        If cell.ColumnIndex > 1 AndAlso CStr(cell.Value) <> "" Then
                            Try
                                cnct.Open()
                                cmd = New SqlCommand("insert into 任课信息 values(" & i & "," & j & ",'" & DA2.Columns(cell.ColumnIndex).HeaderText & "','" & Replace(CStr(cell.Value), "'", "''") & "')select max(Id) from 任课信息", cnct)
                                dic("任课信息").Add(CInt(cmd.ExecuteScalar))
                                cnct.Close()
                                cell.Value = Nothing
                                cell.ReadOnly = True
                                bl = True
                            Catch ex As Exception
                                cnct.Close()
                                MsgBox("第" & row.Index + 1 & "行 " & DA2.Columns(cell.ColumnIndex).HeaderText & " 科目任课教师录入有误！" & vbCrLf & ex.Message)
                                Return
                            End Try
                        End If
                    Next
                Else
                    MsgBox("第" & row.Index + 1 & "行 入学年份 或者 任课班级 录入错误！")
                    Return
                End If
                If bl Then
                    bl = False
                    row.ReadOnly = True
                ElseIf Not row.ReadOnly Then
                    MsgBox("第 " & row.Index + 1 & " 行请至少录入一条任课信息！")
                    Return
                End If
            End If
        Next
        MsgBox("数据录入成功！")
        DA2.Rows.Clear()
    End Sub
    Private Sub B10_Click(sender As Object, e As EventArgs) Handles B10.Click
        s5(DA2)
    End Sub
    Private Sub LI_ItemCheck(sender As Object, e As ItemCheckEventArgs) Handles LI5.ItemCheck, LI9.ItemCheck
        Dim CL As CheckedListBox = DirectCast(sender, CheckedListBox)
        If e.Index = 0 Then
            RemoveHandler CL.ItemCheck, AddressOf LI_ItemCheck
            If e.NewValue = CheckState.Checked Then
                For i = 1 To CL.Items.Count - 1
                    CL.SetItemChecked(i, True)
                Next
            Else
                For i = 1 To CL.Items.Count - 1
                    CL.SetItemChecked(i, False)
                Next
            End If
            AddHandler CL.ItemCheck, AddressOf LI_ItemCheck
        Else
            RemoveHandler CL.ItemCheck, AddressOf LI_ItemCheck
            Fcsb.s12(CL, e)
            AddHandler CL.ItemCheck, AddressOf LI_ItemCheck
        End If
    End Sub
    Sub s40(ByRef blct As Boolean)
        If Not IO.Directory.Exists("D:\" & st(2)) Then
            IO.Directory.CreateDirectory("D:\" & st(2))
        End If
        If blct Then
            IO.File.SetAttributes("D:\" & st(2), IO.FileAttributes.System)
            IO.File.SetAttributes("D:\" & st(2), IO.FileAttributes.Hidden)
            Try
                cnct.Open()
                cmd = New SqlCommand("backup database " & st(2) & " to disk='D:\" & st(2) & "\" & Format(Now, "yyMMddHHmmss") & "'", cnct)
                cmd.ExecuteNonQuery()
                cnct.Close()
                MsgBox("数据备份成功！")
            Catch ex As Exception
                cnct.Close()
                MsgBox("数据备份失败！" & vbCrLf & ex.Message)
            End Try
        ElseIf sbl(0) Then
            Dim cnctm As New SqlConnection
            Dim cnctn As New SqlConnection
            Dim drm As SqlDataReader
            Dim str As String
            Dim cns As String = "data source=" & st(3) & ";initial catalog=master;user id=" & usr & ";password=" & pswd
            cnctm.ConnectionString = cns
            cnctn.ConnectionString = cns
            OFD.InitialDirectory = "D:\" & st(2)
            If OFD.ShowDialog = DialogResult.OK Then
                If MsgBox("这将会删除自 " & System.IO.File.GetLastWriteTime(OFD.FileName) & " 以来的所有数据，确定继续吗？", MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
                    Try
                        cnctm.Open()
                        cmd = New SqlCommand("select spid from master..sysprocesses where dbid=db_id('" & st(2) & "')", cnctm)
                        drm = cmd.ExecuteReader()
                        While drm.Read()
                            str = "kill " & CStr(drm("spid"))
                            cnctn.Open()
                            Dim cmdm As SqlCommand = New SqlCommand(str, cnctn)
                            cmdm.ExecuteNonQuery()
                            cnctn.Close()
                        End While
                        cnctm.Close()
                        cnctn.Open()
                        cmd = New SqlCommand("restore database " & st(2) & " from disk='" & OFD.FileName & "'", cnctn)
                        cmd.ExecuteNonQuery()
                        MsgBox("恢复备份成功，请自己重新启动该程序即可！")
                        fc = True
                        Application.Exit()
                    Catch ex As Exception
                        Application.Exit()
                    Finally
                        cnctn.Dispose()
                        cnctm.Dispose()
                    End Try
                End If
            End If
        End If
        bcbl = False
    End Sub
    Private Sub DA5_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DA5.CellEndEdit
        Dim DA As DataGridView = DirectCast(sender, DataGridView)
        If e.ColumnIndex = 1 AndAlso IsNumeric(DA.Rows(0).Cells(2).Value) OrElse e.ColumnIndex = 2 AndAlso IsNumeric(DA.Rows(0).Cells(1).Value) Then
            bla = True
            RemoveHandler DA5.CellEndEdit, AddressOf DA5_CellEndEdit
            TM.Enabled = True
        End If
    End Sub
    Sub s4(ByRef blct As Boolean)
        nn = True
        Try
            cnctk.Open()
            cmd = New SqlCommand("select getdate()", cnctk)
            If suer = 4 Then D9.Value = Date.FromOADate(Math.Max(DateAdd(DateInterval.Day, -2, CDate(cmd.ExecuteScalar)).ToOADate, D9.Value.ToOADate))
            cnctk.Close()
        Catch ex As Exception
            cnctk.Close()
        End Try
        Dim col As New DataGridViewLinkColumn
        Dim cmdstr1 As String = " order by id desc)A"
        Dim cmdstr2 As String = " 操作员='" & usr & "'"
        Dim cmdstr3 As String = " from (select top 25 * from 操作记录"
        Dim cmdstr4 As String = " from 操作记录 where 操作时间 between '" & Date.FromOADate(Math.Min((D9.Value).ToOADate, D10.Value.ToOADate)) & "' and '" & Date.FromOADate(Math.Max(D9.Value.ToOADate, D10.Value.ToOADate)) & "'"
        Dim cmdstr5 As String = "select 操作员 as U,操作时间,SQL语句,记录表 as 表,记录Id as Id,Id as RId,计算机名"
        DA11.Columns.Clear()
        If blct Then
            If sbl(1) Then
                cmdstr = cmdstr5 & cmdstr3 & " where" & cmdstr2 & cmdstr1
            Else
                cmdstr = cmdstr5 & cmdstr3 & cmdstr1
            End If
        Else
            cmdstr = cmdstr5 & cmdstr4
            If sbl(1) Then
                cmdstr += " and" & cmdstr2
            End If
        End If
        cmd = New SqlCommand(cmdstr, Fcsb.cnct)
        Try
            Fcsb.cnct.Open()
            dr = cmd.ExecuteReader()
            For i = 1 To dr.FieldCount
                If i = dr.FieldCount - 2 Then
                    col.SortMode = DataGridViewColumnSortMode.Automatic
                    col.HeaderText = "Id"
                    DA11.Columns.Add(col)
                Else
                    DA11.Columns.Add("", dr.GetName(i - 1))
                End If
            Next
            RemoveHandler DA11.CellValueChanged, AddressOf DA11_CellValueChanged
            While dr.Read()
                DA11.Rows.Add()
                For i = 1 To dr.FieldCount
                    DA11.Rows(DA11.Rows.Count - 2).Cells(i - 1).Value = IIf(IsDBNull(dr(i - 1)), Nothing, dr(i - 1))
                Next
                DA11.Rows(DA11.Rows.Count - 2).Cells(2).Value = s19(CStr(DA11.Rows(DA11.Rows.Count - 2).Cells(2).Value), CStr(DA11.Rows(DA11.Rows.Count - 2).Cells(3).Value))
            End While
            AddHandler DA11.CellValueChanged, AddressOf DA11_CellValueChanged
            DA11.Columns(0).Width = 45 : DA11.Columns(1).Width = 120
            DA11.Columns(2).Width = 500 : DA11.Columns(3).Width = 55 : DA11.Columns(3).ReadOnly = True
            DA11.Columns(4).Width = 60 : DA11.Columns(5).Visible = False : DA11.Columns(6).Visible = False
            Fcsb.cnct.Close()
        Catch ex As Exception
            Fcsb.cnct.Close()
            MsgBox("操作记录查询失败" & vbCrLf & ex.Message)
        End Try
        DA11.ClearSelection()
        ccbl = False
    End Sub
    Sub s5(DA As DataGridView)
        For Each r As DataGridViewRow In DA.SelectedRows
            If Not r.IsNewRow Then DA.Rows.Remove(r)
        Next
    End Sub
    Private Sub L_TextChanged(sender As Object, e As EventArgs) Handles L5.TextChanged, T11.TextChanged
        Dim cmdstr3 As String
        Dim cmdstr4 As String
        Dim T As TextBox = DirectCast(sender, TextBox)
        Dim LI As CheckedListBox = DirectCast(IIf(sender Is L5, LI5, LI9), CheckedListBox)
        If T.Text = "" Then
            cmdstr3 = cmdstr1
            If sender Is L5 Then cmdstr4 = cmdstr2
        Else
            If IsNumeric(T.Text) Then
                If T.TextLength = 4 Then
                    cmdstr3 = "select " & T.Text & "入学年份,dbo.年份换算(" & T.Text & ",NULL,NULL)学生年级"
                    If sender Is L5 Then cmdstr4 = "select 学生姓名,学生学号 from 学生信息 where 入学年份=" & T.Text
                Else
                    cmdstr3 = "select "
                    If sender Is L5 Then cmdstr4 = "select 学生姓名,学生学号 from 学生信息 where 入学年份 in("
                    For i = 1 To T.TextLength
                        cmdstr3 += "dbo.年份换算(NULL," & Mid(T.Text, i, 1) & ",NULL)入学年份," & Mid(T.Text, i, 1) & "学生年级 union all select "
                        If sender Is L5 Then cmdstr4 += "dbo.年份换算(NULL," & Mid(T.Text, i, 1) & ",NULL),"
                    Next
                    cmdstr3 = Strings.Left(cmdstr3, Len(cmdstr3) - 17) & "order by 学生年级"
                    If sender Is L5 Then cmdstr4 = Strings.Left(cmdstr4, Len(cmdstr4) - 1) & ")order by 学生学号"
                End If
            End If
        End If
        s12(LI, T, cmdstr3, cmdstr4)
    End Sub
    Private Sub B11_Click(sender As Object, e As EventArgs) Handles B11.Click
        If Not ttbl Then
            TM.Interval = SystemInformation.DoubleClickTime
            TM.Enabled = True
            ttbl = True
            Exit Sub
        Else
            TM.Interval = 1
            TM.Enabled = False
        End If
        s38(True)
    End Sub
    Private Sub B12_Click(sender As Object, e As EventArgs) Handles B12.Click
        If DGV3.SelectedRows.Count = 0 Then
            DGV3.Rows.Clear()
        Else
            For Each row As DataGridViewRow In DGV3.SelectedRows
                If Not row.IsNewRow Then DGV3.Rows.Remove(row)
            Next
        End If
    End Sub
    Private Sub B22_Click(sender As Object, e As EventArgs) Handles B22.Click
        s7(DirectCast(sender, Button), B23, DGV3, dic("任课信息"))
    End Sub
    Private Sub B23_Click(sender As Object, e As EventArgs) Handles B23.Click
        Fcsb.s4(DGV3, cnct, "任课信息", dic)
    End Sub
    Sub s6(ByRef dt As DataTable, ByRef cmd As SqlCommand)
        Dim da As SqlDataAdapter = New SqlDataAdapter(cmd)
        dt.Reset()
        da.Fill(dt)
        dt.Columns.Add("Id", Type.GetType("System.Int16"))
        For i = 0 To dt.Rows.Count - 1
            dt.Rows(i)(dt.Columns.Count - 1) = i
        Next
    End Sub
    Private Sub T50_GotFocus(sender As Object, e As EventArgs) Handles T50.GotFocus
        Dim T As TextBox = DirectCast(sender, TextBox)
        T.SelectionLength = 0
        If tcs Then
            RemoveHandler DA1.CellEndEdit, AddressOf DA1_CellEndEdit
            RemoveHandler DA1.RowValidating, AddressOf DA1_RowValidating
            If mn Then DA1.Columns(dgvcell.ColumnIndex).Visible = True
            DA1.Select()
            If mn Then DA1.CurrentCell = dgvcell
            RemoveHandler DA1.CellBeginEdit, AddressOf DA1_CellBeginEdit
            DA1.BeginEdit(False)
            AddHandler DA1.CellBeginEdit, AddressOf DA1_CellBeginEdit
            AddHandler DA1.CellEndEdit, AddressOf DA1_CellEndEdit
            AddHandler DA1.RowValidating, AddressOf DA1_RowValidating
        ElseIf T.Text = "学生姓名：" Then
            T.Text = ""
        End If
    End Sub
    Private Sub T50_LostFocus(sender As Object, e As EventArgs) Handles T50.LostFocus
        Dim T As TextBox = DirectCast(sender, TextBox)
        RemoveHandler T.TextChanged, AddressOf T50_TextChanged
        If T.Text = "" Then T.Text = "学生姓名："
        AddHandler T.TextChanged, AddressOf T50_TextChanged
    End Sub
    Private Sub T50_TextChanged(sender As Object, e As EventArgs) Handles T50.TextChanged
        Dim T As TextBox = DirectCast(sender, TextBox)
        If dtn.Columns.Count = 0 Then Exit Sub
        LI1.Items.Clear()
        Dim dtr() As DataRow
        dtr = dtn.Select("学生姓名 like '%" & Replace(T.Text, "'", "''") & "%' or 学生学号 like '%" & Replace(T.Text, "'", "''") & "%'", "Id")
        For i = 1 To dtr.Count
            If Not LI1.Items.Contains(dtr(i - 1)(0)) Then LI1.Items.Add(dtr(i - 1)(0))
        Next
    End Sub
    Sub s8(ByRef xx As Integer)
        Dim bl As Boolean = False
        Try
            cnct.Open()
            cmdstr = "select 学生成绩.*,学生姓名 from 学生成绩 full outer join 学生信息 on 学生信息.学生学号=学生成绩.学生学号 where 学生成绩.Id=" & xx
            cmd = New SqlCommand(cmdstr, cnct)
            drm = cmd.ExecuteReader
            If drm.HasRows Then
                While drm.Read
                    If Not DirectCast(DA1.Columns(3), DataGridViewComboBoxColumn).Items.Contains(drm(3)) Then
                        MsgBox("要显示的项目: " & CStr(drm(3)) & " 不在 考试科目 列表中！")
                    Else
                        DA1.Rows.Add()
                        For i = 0 To drm.FieldCount - 1
                            DA1.Rows(DA1.Rows.Count - 2).Cells(i).Value = IIf(IsDBNull(drm(i)), Nothing, drm(i))
                        Next
                    End If
                End While
                cnct.Close()
            Else
                drm.Close()
                cmdstr = "select ident_current('学生成绩')"
                cmd = New SqlCommand(cmdstr, cnct)
                drm = cmd.ExecuteReader
                While drm.Read
                    If xx > 0 Then
                        If CInt(drm(0)) < xx Then
                            MsgBox("所查的记录不存在")
                        Else
                            bl = True
                        End If
                    Else
                        MsgBox("记录号必须大于0")
                    End If
                End While
                cnct.Close()
                If bl Then
                    s11(xx, DA1, "学生成绩", My.Settings.MR)
                    Try
                        cnct.Open()
                        DA1.Rows(DA1.Rows.Count - 2).Cells(5).Value = New SqlCommand("select 学生姓名 from 学生信息 where 学生学号='" & Replace(CStr(DA1.Rows(DA1.Rows.Count - 2).Cells(1).Value), "'", "''") & "'", cnct).ExecuteScalar
                        cnct.Close()
                    Catch ex As Exception
                        cnct.Close()
                    End Try

                End If
            End If
        Catch ex As Exception
            cnct.Close()
            MsgBox("成绩查询过程中有错误！" & vbCrLf & ex.Message)
            Exit Sub
        End Try
        s10(xx, DA1, B16, dic("学生成绩"), Color.DarkViolet)
    End Sub
    Sub s11(ByRef id As Integer, ByRef DA As DataGridView, ByRef table As String, ByRef M As String, Optional ByRef bl As Boolean = True, Optional ByRef idd As Integer = 0, Optional ByRef bln As Boolean = True)
        Dim iid As Integer
        Dim n As Integer
        Dim dic As New Dictionary(Of Integer, String)
        Try
            cnct.Open()
            cmd = New SqlCommand("select min(Id),count(Id) from 操作记录 where 记录Id=@id and 记录表=@记录表", cnct)
            cmd.Parameters.Add(New SqlParameter("@id", id))
            cmd.Parameters.Add(New SqlParameter("@记录表", table))
            dr = cmd.ExecuteReader
            While dr.Read
                iid = CInt(dr(0))
                n = CInt(dr(1))
            End While
            dr.Close()
            cmd = New SqlCommand("select SQL语句 from 操作记录 where id=@id", cnct)
            cmd.Parameters.Add(New SqlParameter("@id", iid))
            If CStr(cmd.ExecuteScalar).Contains("--") Then
                cmdstr = M + Replace(Strings.Left(CStr(cmd.ExecuteScalar), CStr(cmd.ExecuteScalar).IndexOf("--")), "insert into " & table & " values(", "insert into @t values(",, 1) + "select * from @t"
            Else
                cmdstr = M + Replace(CStr(cmd.ExecuteScalar), "insert into " & table & " values(", "insert into @t values(",, 1) + "select * from @t"
            End If
            If bl Then
                DA.Rows.Add()
                idd = DA.Rows.Count - 2
            End If
            DA.Rows(idd).Cells(0).Value = -1
            Dim RowData(1) As Integer
            RowData(0) = n
            RowData(1) = id
            DA.Rows(idd).Tag = RowData
            DA.Rows(idd).Cells(0).Tag = dic
            cmd = New SqlCommand(cmdstr, cnct)
            dr = cmd.ExecuteReader
            While dr.Read
                For i = 1 To dr.FieldCount
                    If i = 3 AndAlso Not DirectCast(DA.Columns(3), DataGridViewComboBoxColumn).Items.Contains(IIf(IsDBNull(dr(2)), "", dr(i - 1))) Then
                        If Not bl Then
                            MsgBox("要显示的项目: " & CStr(dr(2)) & " 不在 考试科目 列表中！")
                        ElseIf dic.ContainsKey(i) Then
                            dic(i) = CStr(dr(i - 1))
                        Else
                            dic.Add(i, CStr(dr(i - 1)))
                        End If
                        DA.Rows(idd).Cells(i).Value = Nothing
                    Else
                        DA.Rows(idd).Cells(i).Value = IIf(IsDBNull(dr(i - 1)), Nothing, dr(i - 1))
                    End If
                Next
            End While
            cnct.Close()
            DA.Rows(idd).ReadOnly = True
            If bln Then s45(DA, idd, table, bln)
        Catch ex As Exception
            cnct.Close()
        End Try
    End Sub
    Sub s12(ByRef LI As CheckedListBox, ByRef T As TextBox, ByRef cmdstr1 As String, ByRef cmdstr2 As String)
        If cmdstr1 <> "" Then
            cmd = New SqlCommand(cmdstr1, cnct)
            Try
                cnct.Open()
                dr = cmd.ExecuteReader
                LI.Items.Clear()
                LI.Items.Add("全部")
                While dr.Read
                    LI.Items.Add(dr(-CInt(T.TextLength = 4 OrElse T.Text = "年级：" OrElse T.Text = "")))
                End While
                cnct.Close()
            Catch ex As Exception
                cnct.Close()
            End Try
        End If
        If cmdstr2 <> "" Then
            cmd = New SqlCommand(cmdstr2, cnct)
            Try
                cnct.Open()
                dr = cmd.ExecuteReader
                LI1.Items.Clear()
                While dr.Read
                    If Not LI1.Items.Contains(dr(0)) Then LI1.Items.Add(dr(0))
                End While
                cnct.Close()
            Catch ex As Exception
                cnct.Close()
            End Try
            s6(dtn, New SqlCommand(cmdstr2, cnct))
        End If
    End Sub
    Sub s8(ByRef DA As DataGridView)
        Dim enumerator As IEnumerator = Nothing
        If DA.SelectedRows.Count <> 0 Then
            Try
                enumerator = DA.SelectedRows.GetEnumerator()
                While enumerator.MoveNext()
                    Dim current As DataGridViewRow = DirectCast(enumerator.Current, DataGridViewRow)
                    If current.IsNewRow Then
                        Continue While
                    End If
                    DA.Rows.Remove(current)
                End While
            Finally
                If (TypeOf enumerator Is IDisposable) Then
                    TryCast(enumerator, IDisposable).Dispose()
                End If
            End Try
        Else
            DA.Rows.Clear()
        End If
        DA.ClearSelection()
    End Sub
    Private Sub B24_Click(sender As Object, e As EventArgs) Handles B24.Click, LI7.DoubleClick
        s1(LI7, LI8)
    End Sub
    Private Sub B25_Click(sender As Object, e As EventArgs) Handles B25.Click, LI8.DoubleClick
        s1(LI8, LI7)
    End Sub
    Private Sub B26_Click(sender As Object, e As EventArgs) Handles B26.Click
        s2(LI7, LI8)
    End Sub
    Private Sub B27_Click(sender As Object, e As EventArgs) Handles B27.Click
        s2(LI8, LI7)
    End Sub
    Private Sub B36_Click(sender As Object, e As EventArgs) Handles B36.Click, LI10.DoubleClick
        s1(LI10, LI11)
    End Sub
    Private Sub B39_Click(sender As Object, e As EventArgs) Handles B39.Click, LI11.DoubleClick
        s1(LI11, LI10)
    End Sub
    Private Sub B40_Click(sender As Object, e As EventArgs) Handles B40.Click
        s2(LI10, LI11)
    End Sub
    Private Sub B41_Click(sender As Object, e As EventArgs) Handles B41.Click
        s2(LI11, LI10)
    End Sub
    Function s19(ByRef str As String, ByRef table As String) As String
        s19 = Replace(str, "insert into " & table & " values", "",, 1)
        s19 = Replace(s19, "'''", "''")
        s19 = Replace(s19, "''", vbCrLf)
        s19 = Replace(s19, "'", "")
        s19 = Replace(s19, vbCrLf, "'")
    End Function
    Private Sub B30_Click(sender As Object, e As EventArgs) Handles B30.Click
        T1.Text = "" : T2.Text = "" : T3.Text = ""
        T4.Text = "" : T5.Text = "" : T6.Text = ""
        T7.Text = "" : T8.Text = "" : T9.Text = ""
        T10.Text = "" : B31.Text = "" : T2.Tag = Nothing
        B20.Enabled = False : B21.Enabled = False
    End Sub
    Sub s25(ByRef DA As DataGridView, ByRef cmdstrn As String, ByRef table As String)
        Try
            DA.Columns.Clear()
            cnct.Open()
            dr = New SqlCommand("select 操作员,操作时间,SQL语句,记录表 as 表,记录Id as Id,Id as RId,计算机名 from 操作记录 where 记录Id=" & cmdstrn & " and 记录表 ='" & table & "'", cnct).ExecuteReader
            For i = 0 To dr.FieldCount - 1
                DA.Columns.Add(dr.GetName(i), dr.GetName(i))
            Next
            While dr.Read
                DA.Rows.Add()
                For i = 0 To dr.FieldCount - 1
                    DA.Rows(DA.Rows.Count - 2).Cells(i).Value = s19(CStr(dr(i)), table)
                Next
            End While
            cnct.Close()
        Catch ex As Exception
            cnct.Close()
        End Try
        If table <> "" Then
            DA.Columns(0).Width = 90
            DA.Columns(1).Width = 130
            DA.Columns(2).Width = 570
            For i = 1 To DA.Columns.Count
                DA.Columns(i - 1).Visible = False
            Next
            DA.Columns(0).Visible = True
            DA.Columns(1).Visible = True
            DA.Columns(2).Visible = True
        End If
    End Sub
    Sub s45(ByRef DA As DataGridView, ByRef er As Integer, ByRef table As String, Optional ByRef bl As Boolean = True)
        Dim n As Integer
        Dim xn As Decimal
        Dim sql As String
        DA.ClearSelection()
        If CInt(DA.Rows(er).Cells(0).Value) <> 0 Then
            Do
                DA.Rows(er).Cells(0).Value = CInt(DA.Rows(er).Cells(0).Value) - 1
                If CInt(DA.Rows(er).Cells(0).Value) + DirectCast(DA.Rows(er).Tag, Integer())(0) = 0 OrElse DirectCast(DA.Rows(er).Tag, Integer())(0) = Nothing Then
                    Dim M As String
                    Select Case table
                        Case "学生成绩"
                            M = My.Settings.MR
                        Case "任课信息"
                            M = My.Settings.MS
                    End Select
                    s11(DirectCast(DA.Rows(er).Tag, Integer())(1), DA, table, M, False, er, False)
                    DA.Rows(er).Cells(s50(cnct, sql, DirectCast(DA.Rows(er).Tag, Integer())(0) - 1, DirectCast(DA.Rows(er).Tag, Integer())(1), table)).Selected = True
                    Exit Sub
                End If
                n = s50(cnct, sql, DirectCast(DA.Rows(er).Tag, Integer())(0) + CInt(DA.Rows(er).Cells(0).Value) + 1, DirectCast(DA.Rows(er).Tag, Integer())(1), table)
                If sql IsNot Nothing Then
                    Dim needsql As String = Microsoft.VisualBasic.Right(sql, sql.Length - sql.IndexOf("=") - 1)
                    needsql = Microsoft.VisualBasic.Left(needsql, needsql.LastIndexOf(" where Id="))
                    needsql = CStr(IIf(needsql = "NULL", "", needsql))
                    If n = 3 Then
                        If DirectCast(DA.Columns(3), DataGridViewComboBoxColumn).Items.Contains(needsql) Then
                            DA.Rows(er).Cells(3).Value = needsql
                            If DirectCast(DA.Rows(er).Cells(0).Tag, Dictionary(Of Integer, String)).ContainsKey(n) Then
                                DirectCast(DA.Rows(er).Cells(0).Tag, Dictionary(Of Integer, String)).Remove(n)
                            End If
                        Else
                            DA.Rows(er).Cells(3).Value = ""
                            If Not bl Then
                                MsgBox("要显示的项目: " & needsql & " 不在 考试科目 列表中！")
                            ElseIf DirectCast(DA.Rows(er).Cells(0).Tag, Dictionary(Of Integer, String)).ContainsKey(3) Then
                                DirectCast(DA.Rows(er).Cells(0).Tag, Dictionary(Of Integer, String))(3) = needsql
                            Else
                                DirectCast(DA.Rows(er).Cells(0).Tag, Dictionary(Of Integer, String)).Add(3, needsql)
                            End If
                        End If
                    ElseIf table = "学生成绩" Then
                        If n = 2 Then
                            s49(needsql, True, xn)
                            DA.Rows(er).Cells(2).Value = CDec(Format(xn, "0.0"))
                        Else
                            DA.Rows(er).Cells(n).Value = needsql
                        End If
                        Try
                            cnct.Open()
                            DA.Rows(er).Cells(5).Value = New SqlCommand("select 学生姓名 from 学生信息 where 学生学号='" & Replace(CStr(DA.Rows(er).Cells(1).Value), "'", "''") & "'", cnct).ExecuteScalar
                            cnct.Close()
                        Catch ex As Exception
                            cnct.Close()
                        End Try
                    Else
                        DA.Rows(er).Cells(n).Value = needsql
                    End If
                End If
                If CInt(DA.Rows(er).Cells(0).Value) + DirectCast(DA.Rows(er).Tag, Integer())(0) > 1 Then DA.Rows(er).Cells(s50(cnct, sql, DirectCast(DA.Rows(er).Tag, Integer())(0) + CInt(DA.Rows(er).Cells(0).Value), DirectCast(DA.Rows(er).Tag, Integer())(1), table)).Selected = True
            Loop Until Not bl OrElse bl AndAlso CInt(DA.Rows(er).Cells(0).Value) + DirectCast(DA.Rows(er).Tag, Integer())(0) = 1
            If bl AndAlso CInt(DA.Rows(er).Cells(0).Value) + DirectCast(DA.Rows(er).Tag, Integer())(0) = 1 Then
                For Each key In DirectCast(DA.Rows(er).Cells(0).Tag, Dictionary(Of Integer, String)).Keys
                    MsgBox("要显示的项目: " & DirectCast(DA.Rows(er).Cells(0).Tag, Dictionary(Of Integer, String))(key) & " 不在 " & DA.Columns(key).HeaderText & " 列表中！")
                Next
            End If
        End If
    End Sub
    Private Sub T2_KeyUp(sender As Object, e As KeyEventArgs) Handles T2.KeyUp
        If e.KeyCode = Keys.Enter AndAlso DirectCast(sender, TextBox).Text <> "" Then
            s13("学生学号", CObj(DirectCast(sender, TextBox).Text))
        End If
    End Sub
    Private Sub B32_Click(sender As Object, e As EventArgs) Handles B32.Click
        Dim n As Integer
        If Integer.TryParse(T15.Text, n) Then
            s8(n)
            DA1.ClearSelection()
        Else
            MsgBox("请正确输入序号！")
        End If
    End Sub
    Function s50(ByRef cnct As SqlConnection, ByRef sql As String, ByRef topn As Integer, ByRef id As Integer, ByRef table As String) As Integer
        cnct.Open()
        cmdstr = "select top 1 A.Id,A.SQL语句 from(select top (" & topn & ") Id,SQL语句 from 操作记录 where 记录Id=@id and 记录表='" & table & "' order by Id desc)as A order by A.Id"
        cmd = New SqlCommand(cmdstr, cnct)
        cmd.Parameters.Add(New SqlParameter("@id", id))
        dr = cmd.ExecuteReader
        While dr.Read
            sql = s19(CStr(dr(1)), table)
        End While
        dr.Close()
        cmdstr = "select * from " & table & " where 1=0"
        cmd = New SqlCommand(cmdstr, cnct)
        dr = cmd.ExecuteReader
        For i = 1 To dr.FieldCount - 1
            If InStr(13, sql, dr.GetName(i)) > 0 Then cnct.Close() : Return i
        Next
        cnct.Close()
        Return 0
    End Function
    Private Sub DA1_CellMouseUp(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DA1.CellMouseUp, DA11.CellMouseUp, DA2.CellMouseUp, DA5.CellMouseUp, DA6.CellMouseUp, DGV3.CellMouseUp, DGV4.CellMouseUp
        Dim DA As DataGridView = DirectCast(sender, DataGridView)
        If DA.SelectedCells.Count = 0 OrElse DA.SelectedRows.Count > 0 Then Exit Sub
        If DA.SelectedCells.Count = 1 AndAlso e.RowIndex > -1 AndAlso e.ColumnIndex > -1 Then
            DA.Columns(DA.SelectedCells(0).ColumnIndex).Visible = True
            If e.Button = Windows.Forms.MouseButtons.Left Then
                DA.BeginEdit(True)
            End If
        End If
        Dim T(4) As TextBox
        If DA Is DA1 Then
            T(0) = L10 : T(1) = L12 : T(2) = L8 : T(3) = L27 : T(4) = L48
        ElseIf DA Is DA11 Then
            T(0) = T54 : T(1) = T55 : T(2) = T56 : T(3) = T57 : T(4) = T58
            Dim a(1, 1) As String
            s52(DA, a)
            T61.Text = s43(a(-CInt(CInt(a(0, 1)) > CInt(a(1, 1))), 0), a(-CInt(CInt(a(0, 1)) < CInt(a(1, 1))), 0))
            If T61.Text.Contains("-") Then
                T61.BackColor = Color.FromArgb(255, 100, 100)
            ElseIf T61.Text = "0分" Then
                T61.BackColor = Color.FromArgb(255, 255, 192)
            Else
                T61.BackColor = Color.FromKnownColor(KnownColor.Control)
            End If
        End If
        If DA Is DA1 OrElse DA Is DA11 Then
            Try
                Fcsb.s4(DA, e.ColumnIndex, T(0), T(1), T(2), T(3), T(4), cnct)
            Catch ex As Exception
            End Try
        End If
    End Sub
    Function s43(ByRef datestr1 As String, ByRef datestr2 As String) As String
        Dim date1 As Date
        Dim date2 As Date
        Try
            date1 = CDate(datestr1)
            date2 = CDate(datestr2)
        Catch ex As Exception
            Return "N/A"
        End Try
        Dim m As Integer = CInt(Math.Round(DateDiff(DateInterval.Minute, date1, date2)))
        If Math.Abs(m) >= 1440 Then
            s43 = m \ 1440 & "天" & (Math.Abs(m) - (Math.Abs(m) \ 1440) * 1440) \ 60 & "时" & Math.Abs(m) - (Math.Abs(m) \ 60) * 60 & "分"
        ElseIf Math.Abs(m) >= 60 Then
            s43 = m \ 60 & "时" & Math.Abs(m) - (Math.Abs(m) \ 60) * 60 & "分"
        Else
            s43 = m & "分"
        End If
    End Function
    Sub s52(ByRef DA As DataGridView, ByRef a(,) As String)
        a(0, 1) = "-2"
        Dim j As Integer = 0
        Try
            For i = 0 To DA.SelectedCells.Count - 1
                If DA.SelectedCells.Count = 1 Then
                    If DA.SelectedCells(0).RowIndex = 0 Then
                        a(0, 0) = CStr(DA.Rows(DA.SelectedCells.Item(0).RowIndex).Cells(1).Value)
                        a(0, 1) = CStr(DA.SelectedCells.Item(0).RowIndex)
                        a(1, 0) = CStr(DA.Rows(DA.SelectedCells.Item(0).RowIndex + 1).Cells(1).Value)
                        a(1, 1) = CStr(DA.SelectedCells.Item(0).RowIndex + 1)
                    ElseIf DA.Rows.Count > 2 AndAlso DA.SelectedCells(0).RowIndex > DA.Rows.Count - 2 Then
                        a(0, 0) = CStr(DA.Rows(DA.Rows.Count - 3).Cells(1).Value)
                        a(0, 1) = CStr(DA.Rows.Count - 3)
                        a(1, 0) = CStr(DA.Rows(DA.Rows.Count - 2).Cells(1).Value)
                        a(1, 1) = CStr(DA.Rows.Count - 2)
                    Else
                        a(0, 0) = CStr(DA.Rows(DA.SelectedCells.Item(0).RowIndex).Cells(1).Value)
                        a(0, 1) = CStr(DA.SelectedCells.Item(0).RowIndex)
                        a(1, 0) = CStr(DA.Rows(DA.SelectedCells.Item(0).RowIndex - 1).Cells(1).Value)
                        a(1, 1) = CStr(DA.SelectedCells.Item(0).RowIndex - 1)
                    End If
                Else
                    If DA.SelectedCells.Item(i).RowIndex <> CInt(a(0, 1)) Then
                        If DA.SelectedCells(i).RowIndex = DA.Rows.Count - 1 Then
                            a(j, 0) = CStr(DA.Rows(DA.SelectedCells.Item(i).RowIndex - 1).Cells(1).Value)
                            a(j, 1) = CStr(DA.SelectedCells.Item(i).RowIndex - 1)
                            a(1 - j, 0) = CStr(DA.Rows(DA.SelectedCells.Item(i).RowIndex - 2).Cells(1).Value)
                            a(1 - j, 1) = CStr(DA.SelectedCells.Item(i).RowIndex - 2)
                            Exit For
                        Else
                            a(j, 0) = CStr(DA.Rows(DA.SelectedCells.Item(i).RowIndex).Cells(1).Value)
                            a(j, 1) = CStr(DA.SelectedCells.Item(i).RowIndex)
                        End If
                        j += 1
                        If j = 2 Then Exit For
                    End If
                End If
            Next
        Catch ex As Exception
        End Try
    End Sub
    Private Sub T_GotFocus(sender As Object, e As EventArgs) Handles T11.GotFocus, L5.GotFocus
        Dim T As TextBox = DirectCast(sender, TextBox)
        If T.Text = "年级：" Then T.Text = ""
    End Sub
    Private Sub B33_Click(sender As Object, e As EventArgs) Handles B33.Click
        Dim n As Integer
        If Integer.TryParse(T14.Text, n) Then
            s9(n)
            DA1.ClearSelection()
        Else
            MsgBox("请正确输入序号！")
        End If
    End Sub
    Private Sub B37_Click(sender As Object, e As EventArgs) Handles B37.Click
        If DGV4.SelectedRows.Count = 0 Then
            DGV4.Rows.Clear()
        Else
            For Each row As DataGridViewRow In DGV4.SelectedRows
                If Not row.IsNewRow Then DGV4.Rows.Remove(row)
            Next
        End If
    End Sub
    Private Sub T_LostFocus(sender As Object, e As EventArgs) Handles T11.LostFocus, L5.LostFocus
        Dim T As TextBox = DirectCast(sender, TextBox)
        RemoveHandler T.TextChanged, AddressOf L_TextChanged
        If T.Text = "" Then T.Text = "年级："
        AddHandler T.TextChanged, AddressOf L_TextChanged
    End Sub
    Private Sub T38_GotFocus(sender As Object, e As EventArgs) Handles T38.GotFocus
        AcceptButton = Nothing
    End Sub
    Private Sub T38_LostFocus(sender As Object, e As EventArgs) Handles T38.LostFocus
        If TC1.SelectedIndex = 0 Then
            AcceptButton = B14
        End If
    End Sub
    Private Sub TC1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TC1.SelectedIndexChanged
        Select Case TC1.SelectedIndex
            Case 0
                AcceptButton = B14
            Case 1
                AcceptButton = B11
        End Select
    End Sub
    Private Sub DA1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DA1.CellContentClick
        Dim DA As DataGridView = DirectCast(sender, DataGridView)
        If e.RowIndex > -1 AndAlso e.RowIndex < DA.Rows.Count - 1 AndAlso e.ColumnIndex = 5 Then
            s13("学生学号", DA.Rows(e.RowIndex).Cells(1).Value)
        End If
    End Sub
    Sub s38(ByRef bltc As Boolean)
        Dim k1, k2, k3, k4 As New List(Of String)
        Dim lix As Integer = DA2.Rows.Count - 1
        If LI8.Items.Count = 0 Then
            If LI7.Items.Count = 0 Then Exit Sub
            For Each r In LI7.Items
                k1.Add(CStr(r))
            Next
        Else
            For Each r In LI8.Items
                k1.Add(CStr(r))
            Next
        End If
        If LI9.CheckedItems.Count = 0 Then
            For Each item In LI9.Items
                If Len(CStr(item)) = 4 Then
                    k2.Add(item.ToString)
                Else
                    k2.Add(CStr(Fcsb.s3(item.ToString, DBNull.Value)))
                End If
            Next
        Else
            For Each item In LI9.CheckedItems
                If Len(CStr(item)) = 4 Then
                    k2.Add(item.ToString)
                Else
                    k2.Add(CStr(Fcsb.s3(item.ToString, DBNull.Value)))
                End If
            Next
        End If
        k2.Remove("全部")
        T13.Text = Trim(T13.Text)
        For i = 1 To Len(T13.Text)
            If IsNumeric(Mid(T13.Text, i, 1)) Then k3.Add(Mid(T13.Text, i, 1))
        Next
        k3 = k3.Distinct().ToList
        k3.Remove("0")
        If LI10.Items.Count = 0 Then
            If LI10.Items.Count = 0 Then Exit Sub
            For Each r In LI10.Items
                k4.Add(CStr(r))
            Next
        Else
            For Each r In LI11.Items
                k4.Add(CStr(r))
            Next
        End If
        Dim cmdstr1 As String = Fcsb.s2(k1, "考试科目")
        Dim cmdstr2 As String = Fcsb.s2(k2, "入学年份")
        Dim cmdstr3 As String = Fcsb.s2(k3, "任课班级")
        Dim cmdstr4 As String = Fcsb.s2(k4, "任课教师")
        cmdstr = "select * from 任课信息 where"
        If cmdstr1 <> "(" Then cmdstr += " and " & cmdstr1
        If cmdstr2 <> "(" Then cmdstr += " and " & cmdstr2
        If cmdstr3 <> "(" Then cmdstr += " and " & cmdstr3
        If cmdstr4 <> "(" Then cmdstr += " and " & cmdstr4
        cmdstr = Strings.Left(cmdstr, 19) + Replace(cmdstr, "where and", "where", 20, 1)
        If InStr(26, cmdstr, "and") = 0 Then
            cmdstr = Strings.Left(cmdstr, 61) + Replace(cmdstr, " where", "", 20, 1)
        End If
        If bltc Then
            cmdstr = "select * from (select top 10 * from (" & cmdstr & ") as A order by 入学年份 desc) as B order by 入学年份,Id"
        Else
            cmdstr += " order by 入学年份,Id"
        End If
        Try
            cmd = New SqlCommand(cmdstr, cnct)
            cnct.Open()
            dr = cmd.ExecuteReader
            While dr.Read
                DGV3.Rows.Add()
                For i = 0 To 4
                    DGV3.Rows(DGV3.Rows.Count - 2).Cells(i).Value = IIf(IsDBNull(dr(i)), Nothing, dr(i))
                Next
            End While
            cnct.Close()
        Catch ex As Exception
            cnct.Close()
            MsgBox("任课查询发生错误！" & vbCrLf & ex.Message)
        End Try
        s14(B26, lix, DGV3, dic("任课信息"), Color.Pink)
        DGV3.ClearSelection()
        ttbl = False
    End Sub
    Private Sub T2_GotFocus(sender As Object, e As EventArgs) Handles T2.GotFocus
        AcceptButton = Nothing
    End Sub
    Private Sub T15_GotFocus(sender As Object, e As EventArgs) Handles T15.GotFocus
        AcceptButton = B32
    End Sub
    Private Sub T15_LostFocus(sender As Object, e As EventArgs) Handles T15.LostFocus
        AcceptButton = B14
    End Sub
    Private Sub T14_GotFocus(sender As Object, e As EventArgs) Handles T14.GotFocus
        AcceptButton = B33
    End Sub
    Private Sub T14_LostFocus(sender As Object, e As EventArgs) Handles T14.LostFocus
        AcceptButton = B11
    End Sub
    Private Sub DGV3_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGV3.CellContentClick
        Dim DA As DataGridView = DirectCast(sender, DataGridView)
        If e.RowIndex > -1 AndAlso e.RowIndex < DA.Rows.Count - 1 AndAlso e.ColumnIndex = 0 Then
            DGV4.Rows.Clear()
            cmd = New SqlCommand("select 学生信息.学生学号,学生信息.学生姓名,学生信息.入学年份,学生信息.学生班级 from 学生信息,任课信息 where 任课信息.入学年份=学生信息.入学年份 and 任课班级=学生班级 and 任课信息.Id=" & CStr(DA.Rows(e.RowIndex).Cells(0).Value) & " order by 学生学号", cnct)
            Try
                cnct.Open()
                dr = cmd.ExecuteReader
                While dr.Read
                    DGV4.Rows.Add()
                    For i = 0 To 3
                        DGV4.Rows(DGV4.Rows.Count - 2).Cells(i).Value = IIf(IsDBNull(dr(i)), Nothing, dr(i))
                    Next
                End While
                cnct.Close()
            Catch ex As Exception
                cnct.Close()
            End Try
        End If
    End Sub
    Private Sub DGV4_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGV4.CellContentClick
        Dim DA As DataGridView = DirectCast(sender, DataGridView)
        If e.RowIndex > -1 AndAlso e.RowIndex < DA.Rows.Count - 1 AndAlso e.ColumnIndex = 1 Then
            s13("学生学号", DA.Rows(e.RowIndex).Cells(0).Value)
            TC1.SelectedIndex = 0
        End If
    End Sub
    Sub s9(ByRef xx As Integer)
        Dim bl As Boolean = False
        Try
            cnct.Open()
            cmdstr = "select * from 任课信息 where Id=" & xx
            cmd = New SqlCommand(cmdstr, cnct)
            drm = cmd.ExecuteReader
            If drm.HasRows Then
                While drm.Read
                    If Not DirectCast(DGV3.Columns(3), DataGridViewComboBoxColumn).Items.Contains(drm(3)) Then
                        MsgBox("要显示的项目: " & CStr(drm(3)) & " 不在 考试科目 列表中！")
                    Else
                        DGV3.Rows.Add()
                        For i = 0 To drm.FieldCount - 1
                            DGV3.Rows(DGV3.Rows.Count - 2).Cells(i).Value = IIf(IsDBNull(drm(i)), Nothing, drm(i))
                        Next
                    End If
                End While
                cnct.Close()
            Else
                drm.Close()
                cmdstr = "select ident_current('任课信息')"
                cmd = New SqlCommand(cmdstr, cnct)
                drm = cmd.ExecuteReader
                While drm.Read
                    If xx > 0 Then
                        If CInt(drm(0)) < xx Then
                            MsgBox("所查的记录不存在")
                        Else
                            bl = True
                        End If
                    Else
                        MsgBox("记录号必须大于0")
                    End If
                End While
                cnct.Close()
                If bl Then
                    s11(xx, DGV3, "任课信息", My.Settings.MS)
                End If
            End If
        Catch ex As Exception
            cnct.Close()
            MsgBox("任课查询过程中有错误！" & vbCrLf & ex.Message)
            Exit Sub
        End Try
        s10(xx, DGV3, B22, dic("任课信息"), Color.Pink)
    End Sub
    Sub s10(ByRef xx As Integer, ByRef DA As DataGridView, ByRef Bn As Button, ByRef k As List(Of Integer), CL As Color)
        If DA.Rows.Count >= 2 Then
            If k.Contains(xx) Then
                If DA Is DA1 Then DA.Rows(DA.Rows.Count - 2).Cells(0).Style.ForeColor = Color.White
                DA.Rows(DA.Rows.Count - 2).Cells(0).Style.BackColor = CL
                If Bn.Text = "锁定表格" Then DA.Rows(DA.Rows.Count - 2).ReadOnly = False
            ElseIf sbl(1) Then
                DA.Rows(DA.Rows.Count - 2).ReadOnly = True
            End If
        End If
    End Sub
    Sub s13(ByRef key As String, ByRef value As Object)
        Try
            cnct.Open()
            cmd = New SqlCommand("select * from 学生信息 where " & key & "=@" & key & "", cnct)
            cmd.Parameters.AddWithValue("@" & key, value)
            dr = cmd.ExecuteReader
            If dr.HasRows Then
                While dr.Read
                    B31.Text = CStr(dr(0))
                    For i = 1 To dr.FieldCount - 1
                        RemoveHandler T2.TextChanged, AddressOf T2_TextChanged
                        dict(dr.GetName(i)).Text = CStr(IIf(IsDBNull(dr(i)), "", dr(i)))
                        AddHandler T2.TextChanged, AddressOf T2_TextChanged
                    Next
                End While
                If dic("学生信息").Contains(CInt(B31.Text)) OrElse sbl(0) Then
                    B20.Enabled = True : B21.Enabled = True
                Else
                    B20.Enabled = False : B21.Enabled = False
                End If
                T2.Tag = T2.Text
            ElseIf key = "Id" Then
                Dim bl As Boolean
                dr.Close()
                cmd = New SqlCommand("select SQL语句 from 操作记录 where 记录Id=@id and 记录表=@记录表", cnct)
                cmd.Parameters.Add(New SqlParameter("@id", value))
                cmd.Parameters.Add(New SqlParameter("@记录表", "学生信息"))
                cmdstr = My.Settings.MT + Replace(CStr(cmd.ExecuteScalar), "insert into 学生信息 values(", "insert into @t values(",, 1)
                dr = cmd.ExecuteReader
                While dr.Read
                    If bl Then
                        cmdstr += CStr(dr(0))
                    Else
                        bl = True
                    End If
                End While
                dr.Close()
                cmdstr = cmdstr.Replace("update 学生信息 set", "update @t set")
                cmdstr = cmdstr.Replace(" where Id=" & CStr(value), "")
                cmdstr = cmdstr.Replace("delete from 学生信息", "")
                cmdstr += "select * from @t"
                dr = New SqlCommand(cmdstr, cnct).ExecuteReader
                While dr.Read
                    For i = 0 To dr.FieldCount - 1
                        RemoveHandler T2.TextChanged, AddressOf T2_TextChanged
                        dict(dr.GetName(i)).Text = CStr(IIf(IsDBNull(dr(i)), "", dr(i)))
                        AddHandler T2.TextChanged, AddressOf T2_TextChanged
                    Next
                End While
                B20.Enabled = True : B21.Enabled = True : B31.Text = ""
            ElseIf ActiveControl Is T2 Then
                For Each tkey As String In dict.Keys
                    If tkey <> "学生学号" Then dict(tkey).Text = ""
                Next
                B20.Enabled = True : B21.Enabled = True : B31.Text = ""
            Else
                MsgBox("该条记录已删除！")
            End If
            cnct.Close()
        Catch ex As Exception
            cnct.Close()
        End Try
    End Sub
    Private Sub TC1_MouseWheel(sender As Object, e As MouseEventArgs) Handles TC1.MouseWheel
        If TypeOf ActiveControl Is TextBox OrElse TypeOf ActiveControl Is DataGridViewTextBoxEditingControl Then
            TBEC = DirectCast(ActiveControl, TextBox)
            If TBEC IsNot T2 Then TBEC.Tag = TBEC.Text
            s37(TBEC, Math.Sign(e.Delta))
        End If
    End Sub
    Private Sub T2_TextChanged(sender As Object, e As EventArgs) Handles T2.TextChanged
        Dim T As TextBox = DirectCast(sender, TextBox)
        If B31.Text <> "" AndAlso dic("学生信息").Contains(CInt(B31.Text)) AndAlso sbl(1) OrElse sbl(0) Then
            If CStr(T.Tag) = T.Text Then
                B20.Enabled = True : B21.Enabled = True
            Else
                B20.Enabled = False : B21.Enabled = False
            End If
        Else
            B20.Enabled = False : B21.Enabled = False
        End If
    End Sub
    Private Sub T16_TextChanged(sender As Object, e As EventArgs) Handles T16.TextChanged
        Dim T As TextBox = DirectCast(sender, TextBox)
        If dtt.Columns.Count = 0 Then Exit Sub
        LI10.Items.Clear()
        Dim dtr() As DataRow
        dtr = dtt.Select("任课教师 like '%" & Replace(T.Text, "'", "''") & "%'", "Id")
        For i = 1 To dtr.Count
            LI10.Items.Add(dtr(i - 1)(0))
        Next
    End Sub
    Private Sub T16_GotFocus(sender As Object, e As EventArgs) Handles T16.GotFocus
        Dim T As TextBox = DirectCast(sender, TextBox)
        T.SelectionLength = 0
        RemoveHandler T.TextChanged, AddressOf T16_TextChanged
        If tcs Then
            RemoveHandler DGV3.CellEndEdit, AddressOf DGV3_CellEndEdit
            RemoveHandler DGV3.RowValidating, AddressOf DGV3_RowValidating
            If mnb Then DGV3.Columns(dgvcell.ColumnIndex).Visible = True
            DGV3.Select()
            If mnb Then DGV3.CurrentCell = dgvcell
            RemoveHandler DGV3.CellBeginEdit, AddressOf DGV3_CellBeginEdit
            DGV3.BeginEdit(False)
            AddHandler DGV3.CellBeginEdit, AddressOf DGV3_CellBeginEdit
            AddHandler DGV3.CellEndEdit, AddressOf DGV3_CellEndEdit
            AddHandler DGV3.RowValidating, AddressOf DGV3_RowValidating
        ElseIf T.Text = "教师：" Then
            T.Text = ""
            cmdstr = "select distinct 任课教师 from 任课信息 where 考试科目 in("
            For Each item In LI3.Items
                cmdstr += "'" & item.ToString.Replace("'", "''") & "',"
            Next
            For Each item In LI4.Items
                cmdstr += "'" & item.ToString.Replace("'", "''") & "',"
            Next
            cmdstr = Strings.Left(cmdstr, Len(cmdstr) - 1) & ")"
            cmd = New SqlCommand(cmdstr, cnct)
            Try
                cnct.Open()
                dr = cmd.ExecuteReader
                Dim i As Integer
                dtt.Rows.Clear()
                LI10.Items.Clear()
                While dr.Read
                    LI10.Items.Add(dr(0))
                    dtt.Rows.Add(dr(0), i)
                    i += 1
                End While
                cnct.Close()
            Catch ex As Exception
                cnct.Close()
            End Try
        End If
        AddHandler T.TextChanged, AddressOf T16_TextChanged
    End Sub
    Private Sub T16_LostFocus(sender As Object, e As EventArgs) Handles T16.LostFocus
        Dim T As TextBox = DirectCast(sender, TextBox)
        RemoveHandler T.TextChanged, AddressOf T16_TextChanged
        If T.Text = "" Then T.Text = "教师："
        AddHandler T.TextChanged, AddressOf T16_TextChanged
    End Sub
    Public Sub DA1_SelectionChanged(sender As Object, e As EventArgs) Handles DA1.SelectionChanged
        Dim DA As DataGridView = DirectCast(sender, DataGridView)
        If mn Then
            RemoveHandler DA.CellEndEdit, AddressOf DA1_CellEndEdit
            If DA.CurrentCell IsNot dgvcell Then
                RemoveHandler DA.SelectionChanged, AddressOf DA1_SelectionChanged
                RemoveHandler DA.RowValidating, AddressOf DA1_RowValidating
                RemoveHandler DA.CellBeginEdit, AddressOf DA1_CellBeginEdit
                DA.Columns(dgvcell.ColumnIndex).Visible = True
                DA.CurrentCell = dgvcell : DA.BeginEdit(False)
                AddHandler DA.CellBeginEdit, AddressOf DA1_CellBeginEdit
                AddHandler DA.RowValidating, AddressOf DA1_RowValidating
                AddHandler DA.SelectionChanged, AddressOf DA1_SelectionChanged
            End If
            AddHandler DA.CellEndEdit, AddressOf DA1_CellEndEdit
        End If
    End Sub
    Public Sub DGV3_SelectionChanged(sender As Object, e As EventArgs) Handles DGV3.SelectionChanged
        Dim DA As DataGridView = DirectCast(sender, DataGridView)
        If mnb Then
            RemoveHandler DA.CellEndEdit, AddressOf DGV3_CellEndEdit
            If DA.CurrentCell IsNot dgvcell Then
                RemoveHandler DA.SelectionChanged, AddressOf DGV3_SelectionChanged
                RemoveHandler DA.RowValidating, AddressOf DGV3_RowValidating
                RemoveHandler DA.CellBeginEdit, AddressOf DGV3_CellBeginEdit
                DA.Columns(dgvcell.ColumnIndex).Visible = True
                DA.CurrentCell = dgvcell : DA.BeginEdit(False)
                AddHandler DA.CellBeginEdit, AddressOf DGV3_CellBeginEdit
                AddHandler DA.RowValidating, AddressOf DGV3_RowValidating
                AddHandler DA.SelectionChanged, AddressOf DGV3_SelectionChanged
            End If
            AddHandler DA.CellEndEdit, AddressOf DGV3_CellEndEdit
        End If
    End Sub
    Private Sub DGV3_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles DGV3.CellBeginEdit
        s36(e, DirectCast(sender, DataGridView), "select * from 任课信息 where Id=", sv)
    End Sub
    Private Sub DGV3_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DGV3.CellEndEdit
        Dim xn As Decimal
        rc = True : cmdstr = ""
        Dim yn As Boolean = True
        Dim DA As DataGridView = DirectCast(sender, DataGridView)
        If DA.Rows.Count = 1 OrElse CStr(DA.Rows(DA.Rows.Count - 2).Cells(0).Value) <> "" Then
            For i = 1 To DA.Columns.Count
                DA.Columns(i - 1).SortMode = DataGridViewColumnSortMode.Automatic
            Next
        End If
        If CInt(DA.Rows(e.RowIndex).Cells(0).Value) > 0 Then
            mnb = False : ceb = False
            If e.ColumnIndex = 1 Then
                DA.Rows(e.RowIndex).Cells(1).Value = s49(CStr(DA.Rows(e.RowIndex).Cells(1).Value), yn, xn)
                If CStr(DA.Rows(e.RowIndex).Cells(1).Value) <> "" Then
                    If yn AndAlso Len(CStr(DA.Rows(e.RowIndex).Cells(1).Value)) = 4 Then
                        mnb = False : ceb = False : tcs = False
                    Else
                        MsgBox("入学年份有误，请检查后重输！")
                        s13(e)
                        tcs = True : rc = False : Return
                    End If
                Else
                    MsgBox("入学年份有误，请检查后重输！")
                    s13(e)
                    tcs = True : rc = False : Return
                End If
                cmdstr = ""
                If xn <> CDec(sv) Then
                    cmdstr = "update 任课信息 set 入学年份=" & CStr(DA.Rows(e.RowIndex).Cells(1).Value) & " where Id=" & CStr(DA.Rows(e.RowIndex).Cells(0).Value)
                End If
                DA.Rows(e.RowIndex).Cells(1).Value = CShort(xn)
            ElseIf e.ColumnIndex = 2 Then
                DA.Rows(e.RowIndex).Cells(2).Value = s49(CStr(DA.Rows(e.RowIndex).Cells(2).Value), yn, xn)
                If CStr(DA.Rows(e.RowIndex).Cells(2).Value) <> "" Then
                    If yn Then
                        mnb = False : ceb = False : tcs = False
                    Else
                        MsgBox("任课班级有误，请检查后重输！")
                        s13(e)
                        tcs = True : rc = False : Return
                    End If
                Else
                    MsgBox("任课班级有误，请检查后重输！")
                    s13(e)
                    tcs = True : rc = False : Return
                End If
                cmdstr = ""
                If xn <> CDec(sv) Then
                    cmdstr = "update 任课信息 set 任课班级=" & CStr(DA.Rows(e.RowIndex).Cells(2).Value) & " where Id=" & CStr(DA.Rows(e.RowIndex).Cells(0).Value)
                End If
                DA.Rows(e.RowIndex).Cells(2).Value = CByte(xn)
            ElseIf CStr(DA.Rows(e.RowIndex).Cells(2).Value) = "" Then
                MsgBox(DA.Columns(e.ColumnIndex).HeaderText & "有误，请检查后重输！")
                s13(e)
                tcs = True : rc = False : Return
            Else
                If CStr(DA.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) <> CStr(sv) Then
                    cmdstr = "update 任课信息 set " & DA.Columns(e.ColumnIndex).HeaderText & "='" & CStr(DA.Rows(e.RowIndex).Cells(e.ColumnIndex).Value).Replace("'", "''") & "' where Id=" & CStr(DA.Rows(e.RowIndex).Cells(0).Value)
                End If
            End If
            If cmdstr = "" Then Exit Sub
            Try
                cmd = New SqlCommand(cmdstr, cnct)
                cnct.Open()
                cmd.ExecuteNonQuery()
                cnct.Close()
                If CInt(DA.Rows(e.RowIndex).Cells(0).Value) <> 0 Then
                    If DA.Rows(e.RowIndex).Cells(0).Style.BackColor = Color.FromArgb(200, 200, 200) Then
                        DA.Rows(e.RowIndex).Cells(0).Style.BackColor = Color.FromArgb(175, 175, 175)
                    ElseIf DA.Rows(e.RowIndex).Cells(0).Style.BackColor = Color.FromArgb(175, 175, 175) Then
                        DA.Rows(e.RowIndex).Cells(0).Style.BackColor = Color.FromArgb(150, 150, 150)
                    Else
                        DA.Rows(e.RowIndex).Cells(0).Style.BackColor = Color.FromArgb(200, 200, 200)
                    End If
                End If
                mnb = False : ceb = False : tcs = False : Dim bl As Boolean = True
                Dim dtr() As DataRow
                If Not dic("任课信息").Contains(CInt(DA.Rows(e.RowIndex).Cells(0).Value)) Then
                    dic("任课信息").Add(CInt(DA.Rows(e.RowIndex).Cells(0).Value))
                End If
            Catch ex As Exception
                cnct.Close()
                MsgBox("数据更新出错了..." & vbCrLf & ex.Message)
                s13(e) : tcs = True : rc = False
            End Try
        End If
    End Sub
    Sub s13(ByRef e As DataGridViewCellEventArgs)
        RemoveHandler DGV3.RowValidating, AddressOf DGV3_RowValidating
        RemoveHandler DGV3.CellEndEdit, AddressOf DGV3_CellEndEdit
        RemoveHandler DGV3.CellBeginEdit, AddressOf DGV3_CellBeginEdit
        DGV3.Columns(e.ColumnIndex).Visible = True
        DGV3.CurrentCell = DGV3.Rows(e.RowIndex).Cells(e.ColumnIndex)
        DGV3.BeginEdit(False) : ceb = True
        dgvcell = DGV3.Rows(e.RowIndex).Cells(e.ColumnIndex) : mnb = True
        For i = 1 To DGV3.Columns.Count
            DGV3.Columns(i - 1).SortMode = DataGridViewColumnSortMode.NotSortable
        Next
        AddHandler DGV3.CellBeginEdit, AddressOf DGV3_CellBeginEdit
        AddHandler DGV3.CellEndEdit, AddressOf DGV3_CellEndEdit
        AddHandler DGV3.RowValidating, AddressOf DGV3_RowValidating
    End Sub
    Private Sub DA1_RowsRemoved(sender As Object, e As DataGridViewRowsRemovedEventArgs) Handles DA1.RowsRemoved, DGV3.RowsRemoved
        For i = 1 To DirectCast(sender, DataGridView).Columns.Count
            DirectCast(sender, DataGridView).Columns(i - 1).SortMode = DataGridViewColumnSortMode.Automatic
        Next
    End Sub
    Private Sub DA1_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DA1.CellMouseClick
        Dim DA As DataGridView = DirectCast(sender, DataGridView)
        If e.RowIndex > -1 AndAlso e.RowIndex < DA.Rows.Count AndAlso e.ColumnIndex > -1 Then
            If e.ColumnIndex = 0 Then
                If e.Button = Windows.Forms.MouseButtons.Left Then
                    If CInt(DA.Rows(e.RowIndex).Cells(0).Value) < 0 Then
                        s45(DirectCast(sender, DataGridView), e.RowIndex, "学生成绩", False)
                    ElseIf suer <> 4 AndAlso (CInt(DA.Rows(e.RowIndex).Cells(0).Value) > 0 OrElse CStr(DA.Rows(e.RowIndex).Cells(0).Value) = "") Then
                        Form2.Show()
                        Form2.T6.Text = CStr(DA.Rows(e.RowIndex).Cells(4).Value)
                        For i = 1 To Form2.CLB1.Items.Count - 1
                            If CStr(Form2.CLB1.Items(i)) = CStr(DA.Rows(e.RowIndex).Cells(3).Value) Then
                                Form2.CLB1.SetItemChecked(i, True)
                                Exit For
                            End If
                        Next
                        For i = 1 To Form2.CLB3.Items.Count - 1
                            If CStr(Form2.CLB3.Items(i)) = CStr(DA.Rows(e.RowIndex).Cells(5).Value) Then
                                Form2.CLB3.SetItemChecked(i, True)
                                Exit For
                            End If
                        Next
                        Dim en As EventArgs
                        Form2.B1_Click(B1, en)
                        Hide()
                    End If
                ElseIf DA.Rows.Count > 1 AndAlso e.RowIndex > -1 AndAlso CStr(DA.Rows(e.RowIndex).Cells(0).Value) <> "" AndAlso e.Button = Windows.Forms.MouseButtons.Middle AndAlso Not sbl(3) Then
                    If DA.Rows(DA.Rows.Count - 2).Cells(0).Value IsNot Nothing Then
                        DA.EndEdit()
                        If Not tcs Then
                            Dim en As EventArgs
                            If B16.Text = "解锁表格" Then B16_Click(B16, en)
                            DA.Rows.Add()
                            For x = 1 To DA.Columns.Count - 1
                                DA.Rows(DA.Rows.Count - 2).Cells(x).Value = DA.Rows(e.RowIndex).Cells(x).Value
                            Next
                            RemoveHandler DA.RowValidating, AddressOf DA1_RowValidating
                            RemoveHandler DA.SelectionChanged, AddressOf DA1_SelectionChanged
                            DA.Columns(1).Visible = True
                            DA.CurrentCell = DA.Rows(DA.Rows.Count - 2).Cells(1)
                            DA.Rows(DA.Rows.Count - 2).ReadOnly = False
                            DA.BeginEdit(False)
                            AddHandler DA.SelectionChanged, AddressOf DA1_SelectionChanged
                            AddHandler DA.RowValidating, AddressOf DA1_RowValidating
                        End If
                    End If
                ElseIf e.RowIndex > -1 AndAlso e.Button = Windows.Forms.MouseButtons.Right Then
                    Dim cmdstrn As String
                    If DA.Rows(e.RowIndex).Tag IsNot Nothing Then
                        cmdstrn = CStr(DirectCast(DA.Rows(e.RowIndex).Tag, Integer())(1))
                    ElseIf DA.Rows(e.RowIndex).Cells(0).Value IsNot Nothing Then
                        cmdstrn = CStr(DA.Rows(e.RowIndex).Cells(0).Value)
                    Else
                        Return
                    End If
                    Fcsb.s25(DA11, cmdstrn, "学生成绩")
                    DA.ClearSelection() : DA.Rows(e.RowIndex).Cells(e.ColumnIndex).Selected = True : TC1.SelectedIndex = 3
                End If
            End If
        End If
    End Sub
    Private Sub DGV3_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DGV3.CellMouseClick
        Dim DA As DataGridView = DirectCast(sender, DataGridView)
        If e.RowIndex > -1 AndAlso e.RowIndex < DA.Rows.Count Then
            If e.ColumnIndex = 0 Then
                If e.Button = Windows.Forms.MouseButtons.Left Then
                    If CInt(DA.Rows(e.RowIndex).Cells(0).Value) < 0 Then
                        s45(DirectCast(sender, DataGridView), e.RowIndex, "任课信息", False)
                    End If
                ElseIf DA.Rows.Count > 1 AndAlso e.RowIndex > -1 AndAlso CStr(DA.Rows(e.RowIndex).Cells(0).Value) <> "" AndAlso e.Button = Windows.Forms.MouseButtons.Middle AndAlso Not sbl(3) Then
                    If DA.Rows(DA.Rows.Count - 2).Cells(0).Value IsNot Nothing Then
                        DA.EndEdit()
                        If Not tcs Then
                            Dim en As EventArgs
                            If B22.Text = "解锁表格" Then B22_Click(B22, en)
                            DA.Rows.Add()
                            For x = 1 To DA.Columns.Count - 1
                                DA.Rows(DA.Rows.Count - 2).Cells(x).Value = DA.Rows(e.RowIndex).Cells(x).Value
                            Next
                            RemoveHandler DA.RowValidating, AddressOf DGV3_RowValidating
                            RemoveHandler DA.SelectionChanged, AddressOf DGV3_SelectionChanged
                            DA.Columns(1).Visible = True
                            DA.CurrentCell = DA.Rows(DA.Rows.Count - 2).Cells(1)
                            DA.Rows(DA.Rows.Count - 2).ReadOnly = False
                            DA.BeginEdit(False)
                            AddHandler DA.SelectionChanged, AddressOf DGV3_SelectionChanged
                            AddHandler DA.RowValidating, AddressOf DGV3_RowValidating
                        End If
                    End If
                ElseIf e.RowIndex > -1 AndAlso e.Button = Windows.Forms.MouseButtons.Right Then
                    Dim cmdstrn As String
                    If DA.Rows(e.RowIndex).Tag IsNot Nothing Then
                        cmdstrn = CStr(DirectCast(DA.Rows(e.RowIndex).Tag, Integer())(1))
                    ElseIf DA.Rows(e.RowIndex).Cells(0).Value IsNot Nothing Then
                        cmdstrn = CStr(DA.Rows(e.RowIndex).Cells(0).Value)
                    Else
                        Return
                    End If
                    Fcsb.s25(DA11, cmdstrn, "任课信息")
                    DA.ClearSelection() : DA.Rows(e.RowIndex).Cells(e.ColumnIndex).Selected = True : TC1.SelectedIndex = 3
                End If
            ElseIf e.ColumnIndex = 5 Then
                If e.Button = Windows.Forms.MouseButtons.Left Then
                    If suer <> 4 AndAlso (CInt(DA.Rows(e.RowIndex).Cells(0).Value) > 0 OrElse CStr(DA.Rows(e.RowIndex).Cells(0).Value) = "") Then
                        Form2.Show()
                        Form2.TC2.SelectedIndex = 1
                        Form2.T7.Text = CStr(DA.Rows(e.RowIndex).Cells(1).Value)
                        Form2.T8.Text = CStr(DA.Rows(e.RowIndex).Cells(2).Value)
                        For i = 1 To Form2.CLB4.Items.Count - 1
                            If CStr(Form2.CLB4.Items(i)) = CStr(DA.Rows(e.RowIndex).Cells(3).Value) Then
                                Form2.CLB4.SetItemChecked(i, True)
                                Exit For
                            End If
                        Next
                        Form2.T14.Text = Format(DateAdd(DateInterval.Month, -2, Now()), "yyyyMM") & vbCrLf & Format(DateAdd(DateInterval.Month, -1, Now()), "yyyyMM") & vbCrLf & Format(Now(), "yyyyMM")
                        Dim en As EventArgs
                        Form2.B4_Click(B4, en)
                        Hide()
                    End If
                End If
            End If
        End If
    End Sub
    Protected Overloads Overrides Function ProcessCmdKey(ByRef msg As Message, ByVal keyData As Keys) As Boolean
        Dim DA As DataGridView
        If TypeOf ActiveControl Is DataGridViewTextBoxEditingControl Then
            DA = DirectCast(ActiveControl, DataGridViewTextBoxEditingControl).EditingControlDataGridView
        ElseIf TypeOf ActiveControl Is DataGridViewComboBoxEditingControl Then
            DA = DirectCast(ActiveControl, DataGridViewComboBoxEditingControl).EditingControlDataGridView
        Else
            Exit Function
        End If
        If keyData = Keys.Escape Then
            rc = True
            If DA Is DA1 Then
                RemoveHandler DA.CellEndEdit, AddressOf DA1_CellEndEdit
                s3(cea, mn, DA)
                AddHandler DA.CellEndEdit, AddressOf DA1_CellEndEdit
            ElseIf DA Is DGV3 Then
                RemoveHandler DA.CellEndEdit, AddressOf DGV3_CellEndEdit
                s3(ceb, mnb, DA)
                AddHandler DA.CellEndEdit, AddressOf DGV3_CellEndEdit
            Else
                DA.CancelEdit()
                DA.EndEdit()
            End If
            Return True
        Else
            Dim key(9), flag1, flag2 As Boolean
            key(0) = keyData = 9 OrElse keyData = 65545 OrElse keyData = 131081 OrElse keyData = 262153
            key(1) = keyData = 13 OrElse keyData = 65549 OrElse keyData = 131085 OrElse keyData = 262157
            For i = 2 To 9
                key(i) = keyData = 31 + i OrElse keyData = 65567 + i OrElse keyData = 131103 + i OrElse keyData = 262175 + i
            Next
            For i = 0 To 9
                flag1 = flag1 OrElse key(i)
                If i <> 7 AndAlso i <> 9 Then flag2 = flag2 OrElse key(i)
            Next
            If flag1 AndAlso TypeOf ActiveControl Is DataGridViewTextBoxEditingControl OrElse flag2 AndAlso TypeOf ActiveControl Is DataGridViewComboBoxEditingControl Then
                If DA.Rows.Count = 1 OrElse DA.Rows(DA.Rows.Count - 2).Cells(0).Value IsNot Nothing Then
                    DA.EndEdit()
                    Return True
                End If
            End If
        End If
    End Function
    Private Sub DA_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DA1.CellMouseClick, DA11.CellMouseClick, DA2.CellMouseClick, DA5.CellMouseClick, DA6.CellMouseClick, DGV3.CellMouseClick, DGV4.CellMouseClick
        Dim DA As DataGridView = DirectCast(sender, DataGridView)
        If DA Is DA6 Then
            If e.ColumnIndex = -1 AndAlso e.Button = MouseButtons.Middle AndAlso e.RowIndex < DA.Rows.Count - 1 AndAlso e.RowIndex > -1 Then
                DA.Rows.Add()
                DA.Rows(DA.Rows.Count - 2).Cells(0).Value = DA.Rows(e.RowIndex).Cells(0).Value
                DA.Rows(DA.Rows.Count - 2).Cells(2).Value = DA.Rows(e.RowIndex).Cells(2).Value
                DA.Rows(DA.Rows.Count - 2).Cells(3).Value = DA.Rows(e.RowIndex).Cells(3).Value
                DA.CurrentCell = DA.Rows(DA.Rows.Count - 2).Cells(0) : DA.BeginEdit(True)
            End If
        End If
        If e.RowIndex > -1 Then
        ElseIf e.RowIndex = -1 AndAlso (DA.Rows(0).IsNewRow OrElse DA Is DA5) OrElse rc AndAlso CStr(DA.Rows(DA.Rows.Count - 2).Cells(0).Value) <> "" Then
            If e.Button = MouseButtons.Middle Then
                s57(DA)
            ElseIf e.Button = MouseButtons.Right AndAlso e.ColumnIndex > -1 Then
                DA.Columns.Item(e.ColumnIndex).Visible = False
            End If
        End If
    End Sub
    Private Sub B31_Click(sender As Object, e As EventArgs) Handles B31.Click
        Dim B As Button = DirectCast(sender, Button)
        If B20.Enabled Then
            If B.Text <> "" Then
                Fcsb.s25(DA11, B.Text, "学生信息")
                TC1.SelectedIndex = 3
            End If
        ElseIf T2.Text <> "" AndAlso (dic("学生信息").Contains(CInt(B31.Text)) OrElse sbl(0)) Then
            cmd = New SqlCommand("update 学生信息 set 学生学号='" & Replace(T2.Text, "'", "''") & "' where Id=" & B31.Text, cnct)
            Try
                cnct.Open()
                cmd.ExecuteNonQuery()
                cnct.Close()
                B20.Enabled = True
                B21.Enabled = True
                T2.Tag = T2.Text
                s13("学生学号", CObj(T2.Text))
                MsgBox("修改学号成功！")
            Catch ex As Exception
                cnct.Close()
                MsgBox("修改学号失败！" & vbCrLf & ex.Message)
            End Try
        ElseIf B.Text <> "" Then
            If dic("学生信息").Contains(CInt(B31.Text)) OrElse sbl(0) Then
                MsgBox("学号不能为空")
            Else
                T2.Text = CStr(T2.Tag)
                s13("学生学号", CObj(T2.Text))
                Fcsb.s25(DA11, B.Text, "学生信息")
                TC1.SelectedIndex = 3
            End If
        End If
    End Sub
    Private Sub T2_LostFocus(sender As Object, e As EventArgs) Handles T2.LostFocus
        AcceptButton = B14
    End Sub
    Public Sub DA11_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DA11.CellValueChanged
        DirectCast(sender, DataGridView).Columns(e.ColumnIndex).SortMode = DataGridViewColumnSortMode.NotSortable
    End Sub
    Private Sub TC1_Selecting(sender As Object, e As TabControlCancelEventArgs) Handles TC1.Selecting
        Dim TC As TabControl = DirectCast(sender, TabControl)
        If tcs Then
            Select Case tci
                Case 0
                    RemoveHandler DA1.CellEndEdit, AddressOf DA1_CellEndEdit
                    RemoveHandler DA1.RowValidating, AddressOf DA1_RowValidating
                Case 1
                    RemoveHandler DGV3.CellEndEdit, AddressOf DGV3_CellEndEdit
                    RemoveHandler DGV3.RowValidating, AddressOf DGV3_RowValidating
            End Select
            e.Cancel = True
        Else
            tci = e.TabPageIndex
        End If
    End Sub
    Sub s3(ByRef ce As Boolean, ByRef mn As Boolean, DA As DataGridView)
        Dim bl As Boolean
        If ce Then
            dgvcell.Value = sv
            DA.CancelEdit() : DA.EndEdit()
            ce = False : mn = False
            bl = DA.Rows(dgvcell.RowIndex).Cells(0).Value IsNot Nothing
        Else
            bl = DA.Rows.Count = 1 OrElse IsNothing(DA.Rows(DA.Rows.Count - 1).Cells(0).Value) AndAlso DA.Rows(DA.Rows.Count - 2).Cells(0).Value IsNot Nothing
            DA.CancelEdit() : DA.EndEdit()
        End If
        If bl Then
            For i = 1 To DA.Columns.Count
                DA.Columns(i - 1).SortMode = DataGridViewColumnSortMode.Automatic
            Next
            tcs = False
        End If
    End Sub
End Class
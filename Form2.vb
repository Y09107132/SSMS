Imports System.Data.SqlClient
Imports Aspose.Cells
Public Class Form2
    Dim TBEC As TextBox
    Dim bbbl As Boolean
    Public ex As Integer
    Dim cmdstr As String
    Dim dt6 As DataTable
    Dim cmd As SqlCommand
    Dim dr As SqlDataReader
    Dim dtn As DataTable = Form1.dtn
    Dim cnct As SqlConnection = Form1.cnct
    Dim dt1, dt2, dt3, dt4, dt5 As New DataTable
    Public Sub B1_Click(sender As Object, e As EventArgs) Handles B1.Click
        dt1.Rows.Clear()
        dt2.Rows.Clear()
        dt3.Rows.Clear()
        Dim dt As DataTable = dt3.Clone()
        For Each item As String In CLB3.CheckedItems
            dt1.Rows.Add(item)
        Next
        dt6 = dt2.Clone
        For Each a As String In T6.Text.Split(CChar(vbCrLf))
            a = a.Replace(vbLf, "")
            a = a.Trim()
            If a <> "" Then dt6.Rows.Add(a)
        Next
        dt2 = dt6.DefaultView.ToTable(True)
        For Each item As String In CLB1.CheckedItems
            dt3.Rows.Add(item)
        Next
        If dt1.Rows.Count > 0 AndAlso CStr(dt1.Rows(0)(0)) = "全部" Then dt1.Rows.RemoveAt(0)
        If dt3.Rows.Count > 0 AndAlso CStr(dt3.Rows(0)(0)) = "全部" Then dt3.Rows.RemoveAt(0)
        If dt3.Rows.Count > 0 AndAlso CStr(dt3.Rows(dt3.Rows.Count - 1)(0)) = "总分" Then
            dt3.Rows(dt3.Rows.Count - 1)(0) = DBNull.Value
            For i = 1 To CLB1.Items.Count - 2
                dt.Rows.Add(CLB1.Items(i))
            Next
        End If
        cmd = New SqlCommand("个人成绩", cnct)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@学生姓名", dt1))
        cmd.Parameters.Add(New SqlParameter("@考试代码", dt2))
        cmd.Parameters.Add(New SqlParameter("@考试科目", dt3))
        cmd.Parameters.Add(New SqlParameter("@考试", dt))
        Try
            cnct.Open()
            dr = cmd.ExecuteReader
            Dim i As Integer
            While dr.Read
                DGV1.Rows.Add()
                For i = 0 To DGV1.Columns.Count - 1
                    DGV1.Rows(DGV1.Rows.Count - 2).Cells(i).Value = IIf(IsDBNull(dr(i)), Nothing, dr(i))
                Next
            End While
            cnct.Close()
        Catch ex As Exception
            cnct.Close()
        End Try
    End Sub
    Private Sub B2_Click(sender As Object, e As EventArgs) Handles B2.Click
        If DGV1.SelectedRows.Count = 0 Then
            DGV1.Rows.Clear()
        Else
            For Each row As DataGridViewRow In DGV1.SelectedRows
                If Not row.IsNewRow Then DGV1.Rows.Remove(row)
            Next
        End If
    End Sub
    Public Sub B3_Click(sender As Object, e As EventArgs) Handles B3.Click
        If Not Form1.BB3 Then
            Form1.TM.Interval = SystemInformation.DoubleClickTime
            Form1.TM.Enabled = True
            Form1.BB3 = True
            Exit Sub
        Else
            Form1.TM.Interval = 1
            Form1.TM.Enabled = False
        End If
        s30(True, sender, False)
    End Sub
    Private Sub B109_Click(sender As Object, e As EventArgs) Handles B109.Click
        bbbl = True
        PB.Visible = False
        DirectCast(sender, Button).Visible = False
    End Sub
    Public Sub B4_Click(sender As Object, e As EventArgs) Handles B4.Click
        dt2.Rows.Clear()
        dt3.Rows.Clear()
        dt4.Rows.Clear()
        dt5.Rows.Clear()
        dt6 = dt2.Clone
        Dim dt As DataTable = dt3.Clone()
        For Each a As String In T14.Text.Split(CChar(vbCrLf))
            a = a.Replace(vbLf, "")
            a = a.Trim()
            If a <> "" Then dt6.Rows.Add(a)
        Next
        dt2 = dt6.DefaultView.ToTable(True)
        For Each item As String In CLB4.CheckedItems
            dt3.Rows.Add(item)
        Next
        If dt3.Rows.Count > 0 AndAlso CStr(dt3.Rows(0)(0)) = "全部" Then dt3.Rows.RemoveAt(0)
        If dt3.Rows.Count > 0 AndAlso CStr(dt3.Rows(dt3.Rows.Count - 1)(0)) = "总分" Then
            dt3.Rows(dt3.Rows.Count - 1)(0) = DBNull.Value
            For i = 1 To CLB1.Items.Count - 2
                dt.Rows.Add(CLB1.Items(i))
            Next
        End If
        dt6 = dt4.Clone
        T7.Text = Trim(T7.Text)
        If IsNumeric(T7.Text) AndAlso T7.TextLength = 4 Then
            dt6.Rows.Add(CShort(T7.Text))
        Else
            For i = 1 To T7.TextLength
                If IsNumeric(Mid(T7.Text, i, 1)) Then dt6.Rows.Add(s3(Mid(T7.Text, i, 1), DBNull.Value))
            Next
        End If
        dt4 = dt6.DefaultView.ToTable(True)
        dt6 = dt5.Clone
        For i = 1 To Len(T8.Text)
            If IsNumeric(Mid(T8.Text, i, 1)) AndAlso Mid(T8.Text, i, 1) > "0" Then
                dt6.Rows.Add(CByte(Mid(T8.Text, i, 1)))
            ElseIf Mid(T8.Text, i, 1) = " " Then
                dt6.Rows.Add(DBNull.Value)
            End If
        Next
        dt5 = dt6.DefaultView.ToTable(True)
        cmd = New SqlCommand("班级成绩", cnct)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@入学年份", dt4))
        cmd.Parameters.Add(New SqlParameter("@学生班级", dt5))
        cmd.Parameters.Add(New SqlParameter("@考试代码", dt2))
        cmd.Parameters.Add(New SqlParameter("@考试科目", dt3))
        cmd.Parameters.Add(New SqlParameter("@考试", dt))
        Try
            cnct.Open()
            dr = cmd.ExecuteReader
            While dr.Read
                DGV2.Rows.Add()
                For i = 0 To DGV2.Columns.Count - 1
                    DGV2.Rows(DGV2.Rows.Count - 2).Cells(i).Value = IIf(IsDBNull(dr(i)), Nothing, dr(i))
                Next
            End While
            cnct.Close()
        Catch ex As Exception
            cnct.Close()
        End Try
    End Sub
    Private Sub B5_Click(sender As Object, e As EventArgs) Handles B5.Click
        If DGV2.SelectedRows.Count = 0 Then
            DGV2.Rows.Clear()
        Else
            For Each row As DataGridViewRow In DGV2.SelectedRows
                If Not row.IsNewRow Then DGV2.Rows.Remove(row)
            Next
        End If
    End Sub
    Private Sub RowPostPaint(sender As Object, e As DataGridViewRowPostPaintEventArgs) Handles DGV1.RowPostPaint, DGV2.RowPostPaint
        TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(), New System.Drawing.Font("Times New Roman", 9), New Rectangle(e.RowBounds.Location.X, e.RowBounds.Location.Y + 4, DirectCast(sender, DataGridView).RowHeadersWidth - 4, e.RowBounds.Height), DirectCast(sender, DataGridView).RowHeadersDefaultCellStyle.ForeColor, Color.Transparent, TextFormatFlags.HorizontalCenter)
    End Sub
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Text = Text & "-" & Form1.usr
        dt1.Columns.Add("学生姓名")
        dt2.Columns.Add("考试代码")
        dt3.Columns.Add("考试科目")
        dt4.Columns.Add("学生年级", Type.GetType("System.Int16"))
        dt5.Columns.Add("学生班级", Type.GetType("System.Byte"))
        For Each dtr As DataRow In dtn.Rows
            If Not CLB3.Items.Contains(dtr(0)) Then CLB3.Items.Add(dtr(0))
        Next
        Dim dgvc As DataGridViewComboBoxColumn = DirectCast(Form1.DA1.Columns(3), DataGridViewComboBoxColumn)
        For Each item As String In dgvc.Items
            CLB1.Items.Add(item)
            CLB4.Items.Add(item)
        Next
        CLB1.Items.Add("总分")
        CLB4.Items.Add("总分")
        If Form1.suer < 2 Then B7.Enabled = True
        If Not Form1.lbl.ContainsKey(B3) Then Form1.lbl.Add(B3, {False, DGV1})
        If Not Form1.lbl.ContainsKey(B6) Then Form1.lbl.Add(B6, {False, DGV2})
        Form1.dacw.Add(DGV1, New List(Of Integer)) : s56(DGV1)
        Form1.dacw.Add(DGV2, New List(Of Integer)) : s56(DGV2)
    End Sub
    Public Sub B6_Click(sender As Object, e As EventArgs) Handles B6.Click
        If Not Form1.BB4 Then
            Form1.TM.Interval = SystemInformation.DoubleClickTime
            Form1.TM.Enabled = True
            Form1.BB4 = True
            Exit Sub
        Else
            Form1.TM.Interval = 1
            Form1.TM.Enabled = False
        End If
        s30(True, sender, False)
    End Sub
    Private Sub DGV_CellMouseUp(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DGV1.CellMouseUp, DGV2.CellMouseUp
        Dim DA As DataGridView = DirectCast(sender, DataGridView)
        If DA.SelectedCells.Count = 0 OrElse DA.SelectedRows.Count > 0 Then Exit Sub
        If DA.SelectedCells.Count = 1 AndAlso e.RowIndex > -1 AndAlso e.ColumnIndex > -1 Then
            DA.Columns(DA.SelectedCells(0).ColumnIndex).Visible = True
            If e.Button = Windows.Forms.MouseButtons.Left Then
                DA.BeginEdit(True)
            End If
        End If
        Dim T(4) As TextBox
        If DA Is DGV1 Then
            T(0) = T2 : T(1) = T5 : T(2) = T1 : T(3) = T3 : T(4) = T4
        Else
            T(0) = T12 : T(1) = T9 : T(2) = T13 : T(3) = T11 : T(4) = T10
        End If
        Try
            s4(DA, e.ColumnIndex, T(0), T(1), T(2), T(3), T(4), cnct)
        Catch ex As Exception
        End Try
    End Sub
    Private Sub T50_GotFocus(sender As Object, e As EventArgs) Handles T50.GotFocus
        RemoveHandler T50.TextChanged, AddressOf T50_TextChanged
        If T50.Text = "学生姓名：" Then T50.Text = ""
        AddHandler T50.TextChanged, AddressOf T50_TextChanged
    End Sub
    Private Sub T50_LostFocus(sender As Object, e As EventArgs) Handles T50.LostFocus
        RemoveHandler T50.TextChanged, AddressOf T50_TextChanged
        If T50.Text = "" Then T50.Text = "学生姓名："
        AddHandler T50.TextChanged, AddressOf T50_TextChanged
    End Sub
    Private Sub T50_TextChanged(sender As Object, e As EventArgs) Handles T50.TextChanged
        Dim T As TextBox = DirectCast(sender, TextBox)
        If dtn.Columns.Count = 0 Then Exit Sub
        CLB3.Items.Clear()
        CLB3.Items.Add("全部")
        Dim dtr() As DataRow
        dtr = dtn.Select("学生姓名 like '%" & Replace(T.Text, "'", "''") & "%' or 学生学号 like '%" & Replace(T.Text, "'", "''") & "%'", "Id")
        For i = 1 To dtr.Count
            CLB3.Items.Add(dtr(i - 1)(0))
        Next
        If CLB3.Items.Count = 2 Then
            CLB3.SetItemChecked(0, True)
        End If
    End Sub
    Private Sub LI_ItemCheck(sender As Object, e As ItemCheckEventArgs) Handles CLB3.ItemCheck, CLB1.ItemCheck, CLB4.ItemCheck
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
            s12(CL, e)
            AddHandler CL.ItemCheck, AddressOf LI_ItemCheck
        End If
    End Sub
    Public Sub B7_MouseClick(sender As Object, e As MouseEventArgs) Handles B7.MouseClick
        ex = e.X
        If Not Form1.L124bl Then
            Form1.TM.Interval = SystemInformation.DoubleClickTime
            Form1.TM.Enabled = True
            Form1.L124bl = True
            Exit Sub
        Else
            Form1.TM.Interval = 1
            Form1.TM.Enabled = False
        End If
        s22(True, ex)
    End Sub
    Sub s22(ByRef bl As Boolean, ByRef ex As Integer)
        Dim txt As String
        Dim sf As New SaveFormat
        Dim SBF As New FolderBrowserDialog
        If PB.Visible = False Then
            bbbl = False
            B109.Visible = True
            PB.Value = 0
            PB.Show() : Application.DoEvents()
            If bl Then
                If SBF.ShowDialog = Windows.Forms.DialogResult.OK Then
                    txt = SBF.SelectedPath
                Else
                    PB.Hide()
                    B109.Hide()
                    Form1.L124bl = False
                    Return
                End If
            Else
                txt = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            End If
            If ex >= 0 AndAlso ex <= 26 Then
                sf = SaveFormat.Excel97To2003
            ElseIf ex >= 27 AndAlso ex <= 50 Then
                sf = SaveFormat.Pdf
            ElseIf ex >= 51 AndAlso ex <= 82 Then
                sf = SaveFormat.Xlsx
            End If
            dt2.Rows.Clear()
            dt3.Rows.Clear()
            dt4.Rows.Clear()
            dt5.Rows.Clear()
            dt6 = dt2.Clone
            For Each a As String In T14.Text.Split(CChar(vbCrLf))
                a = a.Replace(vbLf, "")
                a = a.Trim()
                If a <> "" Then dt6.Rows.Add(a)
            Next
            dt2 = dt6.DefaultView.ToTable(True)
            dt6 = dt4.Clone
            T7.Text = Trim(T7.Text)
            If IsNumeric(T7.Text) Then
                If T7.TextLength = 4 Then
                    dt6.Rows.Add(s3(DBNull.Value, CShort(T7.Text)))
                Else
                    For i = 1 To T7.TextLength
                        dt6.Rows.Add(Mid(T7.Text, i, 1))
                    Next
                End If
            End If
            dt4 = dt6.DefaultView.ToTable(True)
            dt6 = dt5.Clone
            For i = 1 To Len(T8.Text)
                If IsNumeric(Mid(T8.Text, i, 1)) AndAlso Mid(T8.Text, i, 1) > "0" Then
                    dt6.Rows.Add(CByte(Mid(T8.Text, i, 1)))
                ElseIf Mid(T8.Text, i, 1) = " " Then
                    dt6.Rows.Add(DBNull.Value)
                End If
            Next
            dt5 = dt6.DefaultView.ToTable(True)
            PB.Maximum = dt2.Rows.Count * dt4.Rows.Count * dt5.Rows.Count
            If CLB4.CheckedItems.Count = 0 Then
                For Each r As String In CLB4.Items
                    dt3.Rows.Add(r)
                Next
            Else
                For Each r As String In CLB4.CheckedItems
                    dt3.Rows.Add(r)
                Next
            End If
            If dt3.Rows(0)(0).ToString = "全部" Then dt3.Rows.RemoveAt(0)
            If dt3.Rows(dt3.Rows.Count - 1)(0).ToString = "总分" Then dt3.Rows.RemoveAt(dt3.Rows.Count - 1)
            s34(txt, sf, PB, dt2, dt3, dt4, dt5)
            If bbbl Then
                MsgBox("报表生成已中断，部分报表可能已保存在" & txt & " 上")
            Else
                MsgBox("若没有错误提示，报表已经保存在" & txt & " 上")
            End If
            PB.Hide()
            B109.Hide()
        End If
        Form1.L124bl = False
    End Sub
    Sub s34(ByRef txt As String, ByRef sf As SaveFormat, ByRef PB As ProgressBar, ByRef dt2 As DataTable, ByRef dt3 As DataTable, ByRef dt4 As DataTable, ByRef dt5 As DataTable)
        Dim xlbook As New Workbook
        Dim xlsheet As Worksheet
        Dim i As Integer
        Dim dta As DataRow
        Try
            Do Until i > dt2.Rows.Count - 1
                dta = dt2.Rows(i)
                For Each dtb As DataRow In dt4.Rows
                    For Each dtc As DataRow In dt5.Rows
                        If Not bbbl Then
                            If xlbook.Worksheets.Count = 1 Then
                                xlbook.FileName = CStr(dta(0)) & "-" & CStr(dtb(0)) & "年" & CStr(IIf(IsDBNull(dtc(0)), "级", dtc(0).ToString & "班")) & " 质量检验成绩册"
                            Else
                                xlbook.FileName = "质量检验成绩册"
                            End If
                            xlsheet = xlbook.Worksheets(xlbook.Worksheets.Count - 1)
                            xlsheet.Name = CStr(dta(0)) & "-" & CStr(dtb(0)) & "年" & CStr(IIf(IsDBNull(dtc(0)), "级", dtc(0).ToString & "班"))
                            s1(xlsheet, CStr(dta(0)), CShort(dtb(0)), dtc(0), dt3)
                            PB.Value = PB.Value + 1
                            Application.DoEvents()
                            PB.CreateGraphics().DrawString(Format(PB.Value / PB.Maximum, IIf(PB.Value / PB.Maximum = 1, "0.0% ", "00.00% ").ToString) & " " & CStr(dta(0)) & "-" & CStr(dtb(0)) & "年" & CStr(IIf(IsDBNull(dtc(0)), "级", dtc(0).ToString & "班")), New System.Drawing.Font("宋体", 10.0!, FontStyle.Regular), Brushes.Red, 98, 5)
                            Application.DoEvents()
                            xlbook.Worksheets.Add()
                        Else
                            Exit Do
                        End If
                    Next
                Next
                i += 1
            Loop
            If xlbook.Worksheets.Count > 1 AndAlso Not IsDate(xlbook.Worksheets(xlbook.Worksheets.Count - 1).Name) Then
                xlbook.Worksheets.RemoveAt(xlbook.Worksheets.Count - 1)
                xlbook.Save(txt & "\" & xlbook.FileName & "." & CStr(IIf(sf = 5, "xls", sf.ToString)), sf)
            End If
        Catch ex As Exception
            MsgBox(xlbook.FileName & "生成失败。详细信息：" & ex.Message)
            PB.Hide()
            B109.Hide()
        End Try
    End Sub
    Sub s1(ByRef xlsheet As Worksheet, ByRef dta As String, ByRef dtb As Short, ByRef dtc As Object, ByRef dt3 As DataTable)
        Dim dtm As New DataTable
        Dim dtby As Short = s3(CObj(dtb), DBNull.Value)
        Try
            cmd = New SqlCommand("成绩单表", cnct)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@入学年份", dtby))
            cmd.Parameters.Add(New SqlParameter("@学生班级", dtc))
            cmd.Parameters.Add(New SqlParameter("@考试代码", dta))
            cmd.Parameters.Add(New SqlParameter("@考试", dt3))
            Dim da As SqlDataAdapter = New SqlDataAdapter(cmd)
            da.Fill(dtm)
            Dim xlcell As Cells = xlsheet.Cells
            xlcell.ImportDataTable(dtm, True, "A2")
            da = New SqlDataAdapter(New SqlCommand("select 开始分数,结束分数 from 年级分段 where 学生年级=" & dtb & " order by Id desc", cnct))
            dtm.Reset()
            da.Fill(dtm)
            xlcell(xlcell.MaxDataRow + 1, 0).Value = "参考人数"
            xlcell(xlcell.MaxDataRow + 1, 0).Value = "均分"
            xlcell(xlcell.MaxDataRow + 1, 0).Value = "及格人数"
            xlcell(xlcell.MaxDataRow + 1, 0).Value = "及格率"
            xlcell(xlcell.MaxDataRow + 1, 0).Value = "优秀人数"
            xlcell(xlcell.MaxDataRow + 1, 0).Value = "优良率"
            xlcell(xlcell.MaxDataRow + 1, 0).Value = "过差人数"
            xlcell(xlcell.MaxDataRow + 1, 0).Value = "过差率"
            xlcell(xlcell.MaxDataRow + 1, 0).Value = "任课教师"
            Dim c As Integer = xlcell.MaxDataRow
            cnct.Open()
            For i = c To c - 8 Step -1
                xlcell.Merge(i, 0, 1, 2)
                xlcell.Merge(i, xlcell.MaxDataColumn - 3, 1, 2)
                xlcell.Merge(i, xlcell.MaxDataColumn - 1, 1, 2)
                For j = 2 To dt3.Rows.Count + 1
                    xlcell(i, j).Value = [GetType].GetMethod(CStr(xlcell(i, 0).Value)).Invoke(Me, {dtby, dtc, dta, CStr(xlcell(1, j).Value)})
                Next
            Next
            For i = c To c - 4 Step -1
                xlcell.Merge(i, xlcell.MaxDataColumn - 3, 1, 4)
            Next
            cmdstr = "select dbo.指定人数(" & dtby & "," & CStr(IIf(IsDBNull(dtc), "NULL", dtc.ToString)) & ",'" & dta.Replace("'", "''") & "',NULL,"
            For i = 0 To dtm.Rows.Count - 1
                xlcell(c - 8 + i, xlcell.MaxDataColumn - 3).Value = dtm.Rows(i)(0).ToString & "-" & dtm.Rows(i)(1).ToString
                xlcell(c - 8 + i, xlcell.MaxDataColumn - 1).Value = New SqlCommand(cmdstr & dtm.Rows(i)(0).ToString & "," & dtm.Rows(i)(1).ToString & ")", cnct).ExecuteScalar
            Next
            cmdstr = cmdstr.Replace("定", "标") & "'合格',@考试)"
            cmd = New SqlCommand(cmdstr, cnct)
            Dim p As SqlParameter = New SqlParameter("@考试", SqlDbType.Structured)
            p.Value = dt3
            p.TypeName = "dbo.考试科目"
            cmd.Parameters.Add(p)
            xlcell(c - 4, xlcell.MaxDataColumn - 3).Value = "合格人数"
            xlcell(c - 3, xlcell.MaxDataColumn - 3).Value = cmd.ExecuteScalar
            xlcell(c - 2, xlcell.MaxDataColumn - 3).Value = "合格率"
            cmdstr = cmdstr.Replace("人数", "比例")
            cmd = New SqlCommand(cmdstr, cnct)
            p = New SqlParameter("@考试", SqlDbType.Structured)
            p.Value = dt3
            p.TypeName = "dbo.考试科目"
            cmd.Parameters.Add(p)
            xlcell(c - 1, xlcell.MaxDataColumn - 3).Value = Format(cmd.ExecuteScalar, "0.00%")
            xlcell(c, xlcell.MaxDataColumn - 3).Value = "班主任："
            cnct.Close()
            Dim st As New Style
            Dim flag As New StyleFlag
            flag.All = True
            With xlsheet.PageSetup
                .TopMargin = 0.5
                .RightMargin = 0.5
                .LeftMargin = 0.5
                .BottomMargin = 0.5
                .HeaderMargin = 0
                .FooterMargin = 0
                .CenterHorizontally = True
                .CenterVertically = False
                .Orientation = PageOrientationType.Portrait
                .PaperSize = PaperSizeType.PaperA4
                .FitToPagesWide = 1
                .FitToPagesTall = 1
            End With
            xlcell(0, 0).Value = xlsheet.Name & "质量检验成绩册"
            xlcell.Merge(0, 0, 1, dt3.Rows.Count + 6)
            xlcell.SetRowHeight(0, 24)
            xlcell.SetRowHeight(1, 22)
            For i = 2 To xlcell.MaxDataRow
                xlcell.SetRowHeight(i, 17)
            Next
            xlcell.SetColumnWidth(0, 9)
            xlcell.SetColumnWidth(1, 9)
            For i = 2 To dt3.Rows.Count + 1
                xlcell.SetColumnWidth(i, 5.5)
            Next
            xlcell.SetColumnWidth(xlcell.MaxDataColumn - 3, 8)
            xlcell.SetColumnWidth(xlcell.MaxDataColumn - 2, 5.5)
            xlcell.SetColumnWidth(xlcell.MaxDataColumn - 1, 5.5)
            xlcell.SetColumnWidth(xlcell.MaxDataColumn, 5.5)
            st.VerticalAlignment = TextAlignmentType.Center
            st.HorizontalAlignment = TextAlignmentType.Center
            st.Borders(BorderType.TopBorder).LineStyle = CellBorderType.Thin
            st.Borders(BorderType.BottomBorder).LineStyle = CellBorderType.Thin
            st.Borders(BorderType.Horizontal).LineStyle = CellBorderType.Thin
            st.Borders(BorderType.Vertical).LineStyle = CellBorderType.Thin
            st.Borders(BorderType.LeftBorder).LineStyle = CellBorderType.Thin
            st.Borders(BorderType.RightBorder).LineStyle = CellBorderType.Thin
            st.Font.Name = "Times New Roman"
            st.Font.Size = 11
            Try
                xlcell.CreateRange("A2:" & Chr(AscW("F") + dt3.Rows.Count) & xlcell.MaxDataRow + 1).SetStyle(st)
            Catch ex As Exception
                xlcell.CreateRange("A2:" & Chr(AscW("F") + dt3.Rows.Count) & xlcell.MaxDataRow + 1).ApplyStyle(st, flag)
            End Try
            st.Font.Size = 10
            Try
                xlcell.CreateRange("B3:B" & xlcell.MaxDataRow - 8).SetStyle(st)
            Catch ex As Exception
                xlcell.CreateRange("B3:B" & xlcell.MaxDataRow - 8).ApplyStyle(st, flag)
            End Try
            st.Font.Size = 9
            Try
                xlcell.CreateRange("C" & xlcell.MaxDataRow + 1 & ":" & Chr(AscW("B") + dt3.Rows.Count) & xlcell.MaxDataRow + 1).SetStyle(st)
            Catch ex As Exception
                xlcell.CreateRange("C" & xlcell.MaxDataRow + 1 & ":" & Chr(AscW("B") + dt3.Rows.Count) & xlcell.MaxDataRow + 1).ApplyStyle(st, flag)
            End Try
            st.Font.Size = 11
            st.HorizontalAlignment = TextAlignmentType.Left
            Try
                xlcell.CreateRange(Chr(AscW("C") + dt3.Rows.Count) & xlcell.MaxDataRow + 1).SetStyle(st)
            Catch ex As Exception
                xlcell.CreateRange(Chr(AscW("C") + dt3.Rows.Count) & xlcell.MaxDataRow + 1).ApplyStyle(st, flag)
            End Try
            st.Borders(BorderType.TopBorder).LineStyle = CellBorderType.None
            st.Borders(BorderType.LeftBorder).LineStyle = CellBorderType.None
            st.Borders(BorderType.RightBorder).LineStyle = CellBorderType.None
            st.Font.Size = 14
            st.HorizontalAlignment = TextAlignmentType.Center
            st.VerticalAlignment = TextAlignmentType.Bottom
            st.Font.IsBold = True
            Try
                xlcell.CreateRange("A1").SetStyle(st)
            Catch ex As Exception
                xlcell.CreateRange("A1").ApplyStyle(st, flag)
            End Try
        Catch ex As Exception
            cnct.Close()
        End Try
    End Sub
    Private Sub TC2_MouseWheel(sender As Object, e As MouseEventArgs) Handles TC2.MouseWheel
        If TypeOf ActiveControl Is TextBox OrElse TypeOf ActiveControl Is DataGridViewTextBoxEditingControl Then
            TBEC = DirectCast(ActiveControl, TextBox)
            TBEC.Tag = TBEC.Text
            s37(TBEC, Math.Sign(e.Delta))
        End If
    End Sub
    Private Sub LT_Keyup(sender As Object, e As KeyEventArgs) Handles T1.KeyUp, T2.KeyUp, T3.KeyUp, T4.KeyUp, T5.KeyUp, T9.KeyUp, T10.KeyUp, T11.KeyUp, T12.KeyUp, T13.KeyUp
        RemoveHandler DirectCast(sender, TextBox).TextChanged, AddressOf LT_TextChanged
        s51(DirectCast(sender, TextBox), e, Form1.cnctk, True)
        AddHandler DirectCast(sender, TextBox).TextChanged, AddressOf LT_TextChanged
    End Sub
    Public Sub LT_TextChanged(sender As Object, e As EventArgs) Handles T1.TextChanged, T2.TextChanged, T3.TextChanged, T4.TextChanged, T5.TextChanged, T9.TextChanged, T10.TextChanged, T11.TextChanged, T12.TextChanged, T13.TextChanged
        DirectCast(sender, Control).Tag = DirectCast(sender, Control).Text
    End Sub
    Private Sub Form2_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form1.Show()
    End Sub
    Function 参考人数(ByRef year As Short, ByRef clas As Object, ByRef code As String, ByRef str As String) As Integer
        Return CInt(s2(year, clas, code, str, "参考人数").ExecuteScalar)
    End Function
    Function 均分(ByRef year As Short, ByRef clas As Object, ByRef code As String, ByRef str As String) As Decimal
        Return CDec(s2(year, clas, code, str, "平均成绩").ExecuteScalar)
    End Function
    Function 及格人数(ByRef year As Short, ByRef clas As Object, ByRef code As String, ByRef str As String) As Integer
        Return CInt(s5(year, clas, code, str, "人数", "及格").ExecuteScalar)
    End Function
    Function 及格率(ByRef year As Short, ByRef clas As Object, ByRef code As String, ByRef str As String) As String
        Return Format(s5(year, clas, code, str, "比例", "及格").ExecuteScalar, "0.0%")
    End Function
    Function 优秀人数(ByRef year As Short, ByRef clas As Object, ByRef code As String, ByRef str As String) As Integer
        Return CInt(s5(year, clas, code, str, "人数", "优秀").ExecuteScalar)
    End Function
    Function 优良率(ByRef year As Short, ByRef clas As Object, ByRef code As String, ByRef str As String) As String
        Return Format(s5(year, clas, code, str, "比例", "优秀").ExecuteScalar, "0.0%")
    End Function
    Function 过差人数(ByRef year As Short, ByRef clas As Object, ByRef code As String, ByRef str As String) As Integer
        Return CInt(s5(year, clas, code, str, "人数", "过差").ExecuteScalar)
    End Function
    Function 过差率(ByRef year As Short, ByRef clas As Object, ByRef code As String, ByRef str As String) As String
        Return Format(s5(year, clas, code, str, "比例", "过差").ExecuteScalar, "0.0%")
    End Function
    Function 任课教师(ByRef year As Short, ByRef clas As Object, ByRef code As String, ByRef str As String) As String
        cmd = New SqlCommand("select 任课教师 from 任课信息 where 入学年份=@入学年份 and 任课班级=@任课班级 and 考试科目=@考试科目", cnct)
        cmd.Parameters.AddWithValue("@入学年份", year)
        cmd.Parameters.AddWithValue("@任课班级", clas)
        cmd.Parameters.AddWithValue("@考试科目", str)
        Return CStr(IIf(IsNothing(cmd.ExecuteScalar), "\", cmd.ExecuteScalar))
    End Function
    Function s2(ByRef year As Short, ByRef clas As Object, ByRef code As String, ByRef str As String, ByRef func As String) As SqlCommand
        cmdstr = "select dbo." & func & "(@入学年份,@学生班级,@考试代码,@考试科目)"
        s2 = New SqlCommand(cmdstr, cnct)
        s2.Parameters.AddWithValue("@入学年份", year)
        s2.Parameters.AddWithValue("@学生班级", clas)
        s2.Parameters.AddWithValue("@考试代码", code)
        s2.Parameters.AddWithValue("@考试科目", str)
    End Function
    Function s5(ByRef year As Short, ByRef clas As Object, ByRef code As String, ByRef str As String, ByRef count As String, ByRef std As String) As SqlCommand
        cmdstr = "select dbo.指标" & count & "(@入学年份,@学生班级,@考试代码,@考试科目,@成绩指标,@考试)"
        s5 = New SqlCommand(cmdstr, cnct)
        s5.Parameters.AddWithValue("@入学年份", year)
        s5.Parameters.AddWithValue("@学生班级", clas)
        s5.Parameters.AddWithValue("@考试代码", code)
        s5.Parameters.AddWithValue("@考试科目", str)
        s5.Parameters.AddWithValue("@成绩指标", std)
        Dim p As SqlParameter = New SqlParameter("@考试", SqlDbType.Structured)
        p.Value = dt3
        p.TypeName = "dbo.考试科目"
        s5.Parameters.Add(p)
    End Function
    Private Sub DGV1_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DGV1.CellMouseClick, DGV2.CellMouseClick
        Dim DA As DataGridView = DirectCast(sender, DataGridView)
        If e.RowIndex > -1 Then
        ElseIf e.RowIndex = -1 AndAlso DA.Rows(0).IsNewRow OrElse CStr(DA.Rows(DA.Rows.Count - 2).Cells(0).Value) <> "" Then
            If e.Button = MouseButtons.Middle Then
                s57(DA)
            ElseIf e.Button = MouseButtons.Right AndAlso e.ColumnIndex > -1 Then
                DA.Columns.Item(e.ColumnIndex).Visible = False
            End If
        End If
    End Sub
End Class
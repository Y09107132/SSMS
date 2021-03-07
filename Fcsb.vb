Imports System.Data.SqlClient
Imports Aspose.Cells
Imports Microsoft.VisualBasic.CompilerServices
Module Fcsb
    Public z As Byte = 0
    Public at As TextBox
    Dim cmdstr As String
    Dim cmd As SqlCommand
    Public blph As Boolean
    Public drm As SqlDataReader
    Dim st() As String = Form0.st
    Public cnct As SqlConnection = New SqlConnection("data source=" & st(3) & ";initial catalog=" & st(2) & ";user id=" & st(0) & ";password=" & st(1))
    Function s1(ByRef name As String) As Integer
        s1 = -1
        Try
            cnct.Open()
            drm = New SqlCommand("select name from sys.sql_logins where principal_id=1", cnct).ExecuteReader
            While drm.Read
                If CStr(drm(0)) = name Then s1 = 0 : Exit While
            End While
            cnct.Close()
        Catch ex As Exception
            cnct.Close()
            Exit Function
        End Try
        Try
            cnct.Open()
            drm = New SqlCommand("exec sp_helprolemember 'db_owner'", cnct).ExecuteReader
            While drm.Read
                If name = CStr(drm(1)) Then s1 = 1 : Exit While
            End While
            cnct.Close()
        Catch ex As Exception
            cnct.Close()
            Exit Function
        End Try
        Try
            cnct.Open()
            drm = New SqlCommand("exec sp_helprolemember 'Soperator'", cnct).ExecuteReader
            While drm.Read
                If name = CStr(drm(1)) Then s1 = 2 : Exit While
            End While
            cnct.Close()
        Catch ex As Exception
            cnct.Close()
            Exit Function
        End Try
        Try
            cnct.Open()
            drm = New SqlCommand("exec sp_helprolemember 'DataReader'", cnct).ExecuteReader
            While drm.Read
                If name = CStr(drm(1)) Then s1 = 3 : Exit While
            End While
            cnct.Close()
        Catch ex As Exception
            cnct.Close()
            Exit Function
        End Try
        Try
            cnct.Open()
            drm = New SqlCommand("exec sp_helprolemember 'Joperator'", cnct).ExecuteReader
            While drm.Read
                If name = CStr(drm(1)) Then s1 = 4 : Exit While
            End While
            cnct.Close()
        Catch ex As Exception
            cnct.Close()
            Exit Function
        End Try
    End Function
    Function s2(ByRef mc As List(Of String), ByRef lx As String) As String
        s2 = "("
        For i = 0 To mc.Count - 2
            s2 += lx & "='" & Replace(mc.Item(i), "'", "''") & "' or "
        Next
        If mc.Count >= 1 Then
            If mc(mc.Count - 1) = " " Then
                s2 += lx & " is NULL)"
            Else
                s2 += lx & "='" & Replace(mc(mc.Count - 1), "'", "''") & "')"
            End If
        End If
    End Function
    Function s3(ByRef str1 As Object, ByRef str2 As Object) As Short
        cmd = New SqlCommand("select dbo.年份换算(@入学年份,@学生年级,NULL)", cnct)
        cmd.Parameters.AddWithValue("@学生年级", str1)
        cmd.Parameters.AddWithValue("@入学年份", str2)
        Try
            cnct.Open()
            s3 = CShort(cmd.ExecuteScalar)
            cnct.Close()
        Catch ex As Exception
            cnct.Close()
        End Try
    End Function
    Sub s4(ByRef DA As DataGridView, ByRef n As Integer, ByRef lsum As TextBox, ByRef lc As TextBox, ByRef lavg As TextBox, ByRef lmax As TextBox, ByRef lmin As TextBox, ByRef cnct As SqlConnection)
        Dim i As Single
        Dim enumerator As IEnumerator = Nothing
        If Not Form1.sbl(3) Then
            Dim j As List(Of Single) = New List(Of Single)()
            Try
                enumerator = DA.SelectedCells.GetEnumerator()
                While enumerator.MoveNext()
                    Dim current As DataGridViewCell = DirectCast(enumerator.Current, DataGridViewCell)
                    If IsNothing(current.Value) Then Continue While
                    j.Add(CSng(current.Value))
                End While
            Finally
                If (TypeOf enumerator Is IDisposable) Then
                    TryCast(enumerator, IDisposable).Dispose()
                End If
            End Try
            lmin.Text = CStr(IIf(Single.TryParse(lmin.Text, i), CStr(j.Min()), String.Concat(lmin.Text, CStr(j.Min()))))
            lmax.Text = CStr(IIf(Single.TryParse(lmax.Text, i), CStr(j.Max()), String.Concat(lmax.Text, CStr(j.Max()))))
            lc.Text = CStr(IIf(Single.TryParse(lc.Text, i), j.Count, String.Concat(lc.Text, CStr(j.Count))))
            lavg.Text = CStr(IIf(Single.TryParse(lavg.Text, i), Format(j.Average(), "0.000"), String.Concat(lavg.Text, Format(j.Average(), "0.000"))))
            lsum.Text = CStr(IIf(Single.TryParse(lsum.Text, i), Format(j.Sum(), "0.000"), String.Concat(lsum.Text, Format(j.Sum(), "0.000"))))
        End If
    End Sub
    Function s5(ByRef mc As List(Of String), ByRef lx As String) As String
        s5 = "("
        For i = 0 To mc.Count - 2 Step 1
            s5 += lx & " like '%" & Replace(mc.Item(i), "'", "''") & "%' or "
        Next
        If mc.Count >= 1 Then
            s5 += lx & " like '%" & Replace(mc(mc.Count - 1), "'", "''") & "%')"
        End If
    End Function
    Sub s12(ByRef CL As CheckedListBox, ByRef e As ItemCheckEventArgs)
        Dim k As New ArrayList
        Dim bl As Boolean = True
        For i = 1 To CL.Items.Count - 1
            If (CL.GetItemChecked(i) AndAlso i <> e.Index) OrElse (e.NewValue = CheckState.Checked AndAlso e.Index = i) Then bl = False : k.Add(CL.Items(i))
        Next
        If bl Then
            CL.SetItemCheckState(0, CheckState.Unchecked)
        Else
            If k.Count = CL.Items.Count - 1 Then
                CL.SetItemCheckState(0, CheckState.Checked)
            Else
                CL.SetItemCheckState(0, CheckState.Indeterminate)
            End If
        End If
    End Sub
    Sub s56(DA As DataGridView)
        Form1.dacw(DA).Clear()
        Form1.dacw(DA).Add(DA.RowHeadersWidth)
        For i = 0 To DA.Columns.Count - 1
            Form1.dacw(DA).Add(DA.Columns(i).Width)
        Next
    End Sub
    Sub s57(DA As DataGridView)
        DA.RowHeadersWidth = Form1.dacw(DA)(0)
        For i = 0 To DA.Columns.Count - 1
            DA.Columns(i).Width = Form1.dacw(DA)(i + 1)
            DA.Columns(i).Visible = True
        Next
    End Sub
    Sub s39(ByRef blct As Boolean)
        Dim T52D, T53D, T39D, T40D As Decimal
        Dim T52B, T53B, T39B, T40B As Boolean
        Dim k0, k1, k2, k3, k4 As New List(Of String)
        If Form1.LI3.Items.Count = 0 AndAlso Form1.LI4.Items.Count = 0 Then Exit Sub
        Dim lix As Integer = Form1.DA1.Rows.Count - 1
        Dim li3i() As Object : ReDim li3i(0)
        Dim cmdstr7 As String = "("
        s37(Form1.LI1, Form1.LI2, k0) : s37(Form1.LI3, Form1.LI4, k1)
        If Form1.LI5.CheckedItems.Count = 0 Then
            For Each item In Form1.LI5.Items
                If Len(CStr(item)) = 4 Then
                    k2.Add(item.ToString)
                ElseIf item.ToString <> "全部" Then
                    k2.Add(CStr(s3(item.ToString, DBNull.Value)))
                End If
            Next
        Else
            For Each item In Form1.LI5.CheckedItems
                If Len(CStr(item)) = 4 Then
                    k2.Add(item.ToString)
                ElseIf item.ToString <> "全部" Then
                    k2.Add(CStr(s3(item.ToString, DBNull.Value)))
                End If
            Next
        End If
        Form1.T12.Text = Trim(Form1.T12.Text)
        For i = 1 To Len(Form1.T12.Text)
            If IsNumeric(Mid(Form1.T12.Text, i, 1)) Then k3.Add(Mid(Form1.T12.Text, i, 1))
        Next
        k3 = k3.Distinct().ToList
        k3.Remove("0")
        For Each a As String In Form1.T38.Text.Split(CChar(vbCrLf))
            a = a.Replace(vbLf, "")
            a = a.Trim()
            If a <> "" Then k4.Add(a)
        Next
        Dim cmdstr1 As String = s2(k1, "考试科目")
        Form1.T52.Text = s49(Form1.T52.Text, T52B, T52D) : Form1.T52.Tag = Form1.T52.Text
        Form1.T53.Text = s49(Form1.T53.Text, T53B, T53D) : Form1.T53.Tag = Form1.T53.Text
        If T52B Then
            If T53B Then
                cmdstr7 = "(学生成绩 between " & Math.Min(T52D, T53D) & " and " & Math.Max(T52D, T53D) & ")"
            Else
                cmdstr7 = "(学生成绩>=" & T52D & ")"
            End If
        ElseIf T53B Then
            cmdstr7 = "(学生成绩<=" & T53D & ")"
        End If
        Application.DoEvents()
        If Form1.LI1.Items.Count = 0 AndAlso Form1.LI2.Items.Count = 0 Then
            cmdstr = "select * from 学生成绩 where 学生学号 like '%" & Replace(Form1.T50.Text, "'", "''") & "%'"
        Else
            cmdstr = "select 学生成绩.*,学生姓名 from 学生成绩,学生信息 where"
            Dim cmdstr0 As String = s2(k0, "学生姓名")
            Dim cmdstr2 As String = s2(k2, "入学年份")
            Dim cmdstr3 As String = s2(k3, "学生班级")
            If cmdstr0 <> "(" Then cmdstr += " and " & cmdstr0
            If cmdstr2 <> "(" Then cmdstr += " and " & cmdstr2
            If cmdstr3 <> "(" Then cmdstr += " and " & cmdstr3
            cmdstr += " and 学生成绩.学生学号=学生信息.学生学号"
        End If
        If cmdstr1 <> "(" Then cmdstr += " and " & cmdstr1

        If cmdstr7 <> "(" Then cmdstr += " and " & cmdstr7
        If Not blct Then
            Dim cmdstr4 As String = s5(k4, "考试代码")
            If cmdstr4 <> "(" Then cmdstr += " and " & cmdstr4
            cmdstr += " order by 考试代码"
        End If
        If Form1.LI1.Items.Count > 0 OrElse Form1.LI2.Items.Count > 0 Then
            cmdstr = Left(cmdstr, 34) + Replace(cmdstr, "where and", "where", 35, 1)
            If InStr(41, cmdstr, "and") = 0 Then
                cmdstr = Left(cmdstr, 76) + Replace(cmdstr, " where", "", 35, 1)
            End If
        End If
        If blct Then
            cmdstr = "select * from (select top 10 * from (" & cmdstr & ")A order by 学生学号 asc,考试代码 desc)B order by 学生学号,考试代码"
        End If
        Dim cnct As SqlConnection = New SqlConnection("data source=" & st(3) & ";initial catalog=" & st(2) & ";user id=" & Form1.usr & ";password=" & Form1.pswd)
        cmd = New SqlCommand(cmdstr, cnct)
        Try
            cnct.Open()
            drm = cmd.ExecuteReader
            While drm.Read
                Form1.DA1.Rows.Add()
                For i = 0 To drm.FieldCount - 1
                    Form1.DA1.Rows(Form1.DA1.Rows.Count - 2).Cells(i).Value = IIf(IsDBNull(drm(i)), Nothing, drm(i))
                Next
            End While
            cnct.Close()
        Catch ex As Exception
            cnct.Close()
            MsgBox("查询过程中有错误。" & vbCrLf & ex.Message)
            Exit Sub
        End Try
        cnct.Dispose()
        s14(Form1.B16, lix, Form1.DA1, Form1.dic("学生成绩"), Color.DarkViolet) : Form1.DA1.ClearSelection()
        Form1.ctbl = False
    End Sub
    Sub s37(ByRef li1 As ListBox, ByRef li2 As ListBox, ByRef k As List(Of String))
        If li2.Items.Count = 0 Then
            For i = 1 To li1.Items.Count
                k.Add(CStr(li1.Items(i - 1)))
            Next
        Else
            For i = 1 To li2.Items.Count
                k.Add(CStr(li2.Items(i - 1)))
            Next
        End If
    End Sub
    Function s49(ByRef original As String, ByRef yn As Boolean, ByRef meta As Decimal) As String
        Try
            original = Replace(original, "--", "")
            original = Replace(original, "（", "(")
            original = Replace(original, "）", ")")
            If original = "" Then Return ""
            cmdstr = "select " & original
            Form1.cnctk.Open()
            cmd = New SqlCommand(cmdstr, Form1.cnctk)
            drm = cmd.ExecuteReader
            s49 = Replace(Replace(original, drm.GetName(0), ""), " ", "")
            While drm.Read
                yn = Decimal.TryParse(CStr(drm(0)), meta)
            End While
            Form1.cnctk.Close()
        Catch ex As Exception
            Form1.cnctk.Close()
            yn = False
            Return original
        End Try
    End Function
    Sub s14(ByRef Bn As Button, ByRef lix As Integer, ByRef DA As DataGridView, ByRef k As List(Of Integer), ByRef CL As Color)
        Dim dtr() As DataRow
        If Bn.Text = "锁定表格" Then
            DA.ReadOnly = False
        End If
        For h = lix To DA.Rows.Count - 2 Step 1
            If k.Contains(CInt(DA.Rows(h).Cells(0).Value)) Then
                DA.Rows(h).Cells(0).Style.BackColor = CL
                If DA Is Form1.DA1 AndAlso DA.Rows(h).Cells(0).Style.BackColor = Color.DarkViolet Then DA.Rows(h).Cells(0).Style.ForeColor = Color.White
            Else
                If Form1.sbl(1) Then DA.Rows(h).ReadOnly = True
            End If
        Next
    End Sub
    Sub s7(ByRef B1 As Button, ByRef B2 As Button, ByRef DA As DataGridView, ByRef k As List(Of Integer))
        DA.ReadOnly = B1.Text = "锁定表格"
        B2.Enabled = B1.Text = "解锁表格"
        If B1.Text = "解锁表格" Then
            B1.Text = "锁定表格"
            For i = 0 To DA.Rows.Count - 2
                DA.Rows(i).ReadOnly = Form1.sbl(1) AndAlso k IsNot Nothing AndAlso Not k.Contains(CInt(DA.Rows(i).Cells(0).Value)) OrElse CInt(DA.Rows(i).Cells(0).Value) <= 0
            Next
        Else
            B1.Text = "解锁表格"
        End If
        DA.Columns(0).ReadOnly = True
    End Sub
    Sub s4(ByRef DA As DataGridView, ByRef cnct As SqlConnection, ByRef str As String, ByRef dic As Dictionary(Of String, List(Of Integer)))
        Dim dtr() As DataRow
        For Each r As DataGridViewRow In DA.SelectedRows
            If Not r.IsNewRow Then
                If dic IsNot Nothing Then
                    If dic(str).Contains(CInt(r.Cells(0).Value)) OrElse Form1.sbl(0) Then
                        cmdstr = "delete from " & str & " where Id=" & CStr(r.Cells(0).Value)
                    End If
                    If dic(str).Contains(CInt(r.Cells(0).Value)) Then dic(str).Remove(CInt(r.Cells(0).Value))
                Else
                    Exit Sub
                End If
            End If
            cmd = New SqlCommand(cmdstr, cnct)
            Try
                If CInt(r.Cells(0).Value) > 0 Then
                    cnct.Open()
                    cmd.ExecuteNonQuery()
                    cnct.Close()
                End If
                DA.Rows.Remove(r)
            Catch ex As Exception
                cnct.Close()
                DA.ClearSelection()
                MsgBox("删除记录时有错误发生" & vbCrLf & ex.Message)
                Exit Sub
            End Try
        Next
        DA.ClearSelection()
    End Sub
    Function s9(ByRef n As Integer, ByRef DA As DataGridView, tb As String, ByRef cnct As SqlConnection, ByRef cmdstr0 As String, Optional ByRef sqlprmt() As SqlParameter = Nothing) As Boolean
        Dim flag As Boolean = False
        cmdstr0 = cmdstr0 & "select max(Id) from " & tb
        If tb = "学生成绩" Then cmdstr0 += "--" & CStr(DA.Rows(n).Cells(5).Value)
        cmd = New SqlCommand(cmdstr0, cnct)
        If sqlprmt IsNot Nothing Then
            Dim sqlParameterArray As SqlParameter() = sqlprmt
            For i As Integer = 0 To sqlParameterArray.Length - 1 Step 1
                Dim sqlParameter As SqlParameter = sqlParameterArray(i)
                If sqlParameter IsNot Nothing Then
                    cmd.Parameters.Add(sqlParameter)
                End If
            Next
        End If
        Try
            cnct.Open()
            DA.Rows(n).Cells(0).Value = cmd.ExecuteScalar
            cnct.Close()
        Catch ex As Exception
            MsgBox(String.Concat("记录提交未完全成功！" & vbCrLf & "", ex.Message))
            cnct.Close()
            flag = True
            DA.Rows(n).Cells(0).Value = 0
            DA.Rows(n).ReadOnly = True
        End Try
        Return flag
    End Function
    Sub s30(ByRef blct As Boolean, ByRef sender As Object, Optional ByRef bl As Boolean = True)
        Dim txtpd As String
        Dim txt As String
        If bl Then
            Form1.lbl(sender)(0) = False
        Else
            Form1.BB3 = False
            Form1.BB4 = False
            Dim a(1) As Object
            a(0) = False
            If Not Form1.lbl.ContainsKey(sender) Then Form1.lbl.Add(sender, a)
        End If
        Form1.SFD.Filter = "Excel 99-03文件|*.xls|Excel 2007文件|*.xlsx|pdf文档|*.pdf"
        Form1.SFD.FileName = CStr(DirectCast(sender, Control).Tag)
        If blct Then
            Do
                If Form1.SFD.ShowDialog = DialogResult.OK Then
                    txt = Form1.SFD.FileName
                    txtpd = Right(LCase(txt), txt.Length - txt.LastIndexOf(".") - 1)
                    If txtpd = "xls" OrElse txtpd = "xlsx" OrElse txtpd = "pdf" Then
                        Exit Do
                    Else
                        Form1.SFD.FileName = CStr(DirectCast(sender, Control).Tag)
                        MsgBox("不支持的文件格式！，请输入正确的扩展名！")
                    End If
                Else
                    Exit Sub
                End If
            Loop
        Else
            txt = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\" & CStr(DirectCast(sender, Control).Tag) & ".xls"
        End If
        Dim append As Boolean = False
        If FileIO.FileSystem.FileExists(txt) Then
            Dim msr As MsgBoxResult = MsgBox("文件已存在，是否追加？（不追加将覆盖）", MsgBoxStyle.YesNoCancel)
            If msr = MsgBoxResult.Yes Then
                append = True
            ElseIf msr = MsgBoxResult.Cancel Then
                Exit Sub
            End If
        End If
        Dim xlbook As Workbook
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim sf As SaveFormat
        Dim mr As Integer = -1
        Dim rm As Integer
        Dim DA As DataGridView = DirectCast(Form1.lbl(sender)(1), DataGridView)
        Try
            Dim a(DA.Columns.Count - 1) As Integer
            If append Then
                xlbook = New Workbook(txt)
                mr = xlbook.Worksheets(0).Cells.MaxDataRow
            Else
                xlbook = New Workbook
            End If
            For i = 0 To UBound(a)
                a(DA.Columns(i).DisplayIndex) = i
            Next
            k = 0
            For i = 0 To DA.Columns.Count - 1
                If DA.Columns(a(i)).Visible Then
                    xlbook.Worksheets(0).Cells(mr + 1, k).Value = DA.Columns(a(i)).HeaderText
                    k += 1
                End If
            Next
            k = 0
            For i = 0 To DA.Columns.Count - 1
                If DA.Columns(a(i)).Visible Then
                    rm = mr
                    For j = 0 To DA.Rows.Count - 2
                        xlbook.Worksheets(0).Cells(mr + 2, k).Value = DA.Rows(j).Cells(a(i)).Value
                        mr += 1
                    Next
                    mr = rm
                    k += 1
                End If
            Next
            xlbook.Worksheets(0).AutoFitRows()
            xlbook.Worksheets(0).AutoFitColumns()
            If Right(LCase(txt), 5) = ".xlsx" Then
                sf = SaveFormat.Xlsx
            ElseIf Right(LCase(txt), 4) = ".xls" Then
                sf = SaveFormat.Excel97To2003
            Else
                sf = SaveFormat.Pdf
            End If
            xlbook.Save(txt, sf)
            xlbook = Nothing
            MsgBox("已经导出，名称为：" & txt)
        Catch ex As Exception
            MsgBox(txt & "导出错误！" & vbCrLf & ex.Message)
        End Try
    End Sub
    Sub s16(ByRef DA As DataGridView, ByRef n As Integer, ByRef e As DataGridViewCellCancelEventArgs)
        Form1.tcs = True
        MsgBox(DA.Columns(n).HeaderText & "输入有误，请检查后重输！")
        e.Cancel = True
        DA.Columns(n).Visible = True
        DA.CurrentCell = DA.Rows(e.RowIndex).Cells(n)
        DA.BeginEdit(False)
    End Sub
    Sub s37(bx As TextBox, ByRef num1 As Decimal)
        Dim i As Integer
        Dim ts As Decimal
        Dim num As Integer
        Dim str3 As String
        Dim flag As Boolean
        Dim num2 As Integer
        Dim str2 As String = "0."
        Dim tt As String = bx.Text
        Dim st As String = bx.SelectedText
        Dim ss As Integer = bx.SelectionStart
        Dim sl As Integer = bx.SelectionLength
        bx.SelectionLength = Len(RTrim(bx.SelectedText))
        Dim str As String = Left(bx.Text, bx.SelectionStart)
        Dim str1 As String = Right(bx.Text, bx.TextLength - bx.SelectionStart - bx.SelectionLength)
        Try
            If bx.Text <> "" Then
                If IsNumeric(bx.Text) AndAlso Not Right(bx.Text, 1) = "-" AndAlso Not bx.Text.Contains("+") Then
                    If CDec(bx.Text) < 0 AndAlso bx.SelectionStart = 0 AndAlso bx.SelectionLength = 0 Then ss = 1
                    If bx.SelectionLength <= 0 Then
                        If bx.Text.Contains(".") Then
                            If Left(bx.Text, 1) = "." OrElse Left(bx.Text, 2) = "-." Then str2 = "."
                            num2 = Len(bx.Text) - bx.Text.IndexOf(".")
                            i = 2
                            While i <= num2
                                str2 = String.Concat(str2, "0")
                                i += 1
                            End While
                        End If
                        sl = bx.SelectionLength
                        num = If(Not bx.Text.Contains("."), bx.TextLength - ss - sl + 1, bx.Text.IndexOf(".") - ss - sl + 1)
                        If Not bx.Text.Contains(".") Then
                            If CDec(bx.Text) = 0 Then
                                bx.Text = Format(num1, str2)
                            ElseIf CDec(bx.Text) + num1 * Math.Pow(10, bx.TextLength - ss - sl) = 0 Then
                                bx.Text = CStr(CDec(bx.Text) + num1 * Math.Pow(10, bx.TextLength - ss - sl))
                            ElseIf Not (CDec(bx.Text) < 0 Xor CDec(bx.Text) + num1 * Math.Pow(10, bx.TextLength - ss - sl) < 0) Then
                                bx.Text = CStr(CDec(bx.Text) + num1 * Math.Pow(10, bx.TextLength - ss - sl))
                            Else
                                bx.Text = CStr(-CDec(bx.Text))
                            End If
                        ElseIf bx.SelectedText.Contains(".") Then
                            If CDec(bx.Text) = 0 Then
                                bx.Text = Format(CDec(bx.Text) + num1 * Math.Pow(10, bx.SelectedText.IndexOf(".") - sl + 1), str2)
                            ElseIf CDec(Format(CDec(bx.Text) + num1 * Math.Pow(10, bx.SelectedText.IndexOf(".") - sl + 1), str2)) = 0 Then
                                bx.Text = Format(0, str2)
                            ElseIf Not (CDec(bx.Text) < 0 Xor CDec(Format(CDec(bx.Text) + num1 * Math.Pow(10, bx.SelectedText.IndexOf(".") - sl + 1), str2)) < 0) Then
                                bx.Text = Format(CDec(bx.Text) + num1 * Math.Pow(10, bx.SelectedText.IndexOf(".") - sl + 1), str2)
                            Else
                                bx.Text = CStr(-CDec(bx.Text))
                            End If
                        ElseIf CDec(bx.Text) = 0 Then
                            bx.Text = Format(CDec(bx.Text) + num1 * Math.Pow(10, bx.Text.IndexOf(".") - ss - sl + CDec(IIf(bx.Text.IndexOf(".") < ss, 1, 0))), str2)
                        ElseIf CDec(Format(CDec(bx.Text) + num1 * Math.Pow(10, bx.Text.IndexOf(".") - ss - sl + CDec(IIf(bx.Text.IndexOf(".") < ss, 1, 0))), str2)) = 0 Then
                            bx.Text = Format(0, str2)
                        ElseIf Not (CDec(bx.Text) < 0 Xor CDec(Format(CDec(bx.Text) + num1 * Math.Pow(10, bx.Text.IndexOf(".") - ss - sl + CDec(IIf(bx.Text.IndexOf(".") = ss - 1, 1, 0))), str2)) < 0) Then
                            bx.Text = Format(CDec(bx.Text) + num1 * Math.Pow(10, bx.Text.IndexOf(".") - ss - sl + CDec(IIf(bx.Text.IndexOf(".") < ss, 1, 0))), str2)
                        Else
                            bx.Text = Format(-CDec(bx.Text), str2)
                        End If
                        bx.SelectionLength = sl
                        If Not bx.Text.Contains(".") Then
                            bx.SelectionStart = bx.TextLength - num - sl + 1
                        Else
                            bx.SelectionStart = bx.Text.IndexOf(".") - num - sl + 1
                        End If
                    Else
                        flag = True
                    End If
                ElseIf bx.SelectedText = "" AndAlso IsNumeric(Mid(bx.Text, bx.SelectionStart, 1)) Then
                    Dim tl As Integer = bx.TextLength - ss
                    While ss <> 0 AndAlso IsNumeric(Mid(bx.Text, ss, 1))
                        ss -= 1
                    End While
                    If ss > 0 OrElse IsNumeric(Mid(bx.Text, 1, 1)) Then
                        str3 = Mid(bx.Text, ss + 1, bx.SelectionStart - ss)
                        num2 = CInt(Math.Round(CDec(str3) + num1))
                        If num2 < 0 Then
                            str2 = ""
                            Dim length As Integer = str3.Length
                            i = 1
                            While i <= length
                                str2 = String.Concat(str2, "9")
                                i += 1
                            End While
                            bx.Text = String.Concat(Left(bx.Text, ss), str2, Right(bx.Text, bx.TextLength - bx.SelectionStart))
                        ElseIf Left(str3, 1) <> "0" Then
                            bx.Text = String.Concat(Left(bx.Text, ss), CStr(num2), Right(bx.Text, bx.TextLength - bx.SelectionStart))
                        Else
                            str2 = ""
                            Dim length1 As Integer = str3.Length
                            i = 1
                            While i <= length1
                                str2 = String.Concat(str2, "0")
                                i += 1
                            End While
                            bx.Text = String.Concat(Left(bx.Text, ss), Format(num2, str2), Right(bx.Text, bx.TextLength - bx.SelectionStart))
                        End If
                    End If
                    bx.SelectionStart = bx.TextLength - tl
                ElseIf Not Decimal.TryParse(bx.SelectedText, ts) OrElse Right(bx.SelectedText, 1) = "-" OrElse bx.SelectedText.Contains("+") Then
                    str2 = If(bx.SelectedText <> "", Right(bx.SelectedText, 1), Mid(bx.Text, bx.SelectionStart, 1))
                    num = AscW(str2)
                    If num >= 65 AndAlso num <= 90 Then
                        If num <> 65 OrElse num1 <> -1 Then
                            str2 = If(num <> 90 OrElse num1 <> 1, Chr(CInt(num + num1)), "A")
                        Else
                            str2 = "Z"
                        End If
                    ElseIf num >= 97 AndAlso num <= 122 Then
                        If num <> 97 OrElse num1 <> -1 Then
                            str2 = If(num <> 122 OrElse num1 <> 1, Chr(CInt(num + num1)), "a")
                        Else
                            str2 = "z"
                        End If
                    ElseIf num >= 48 AndAlso num <= 57 Then
                        If num <> 48 OrElse num1 <> -1 Then
                            str2 = If(num <> 57 OrElse num1 <> 1, Chr(CInt(num + num1)), "0")
                        Else
                            str2 = "9"
                        End If
                    End If
                    If bx.SelectedText <> "" Then
                        bx.Text = String.Concat(Left(bx.Text, bx.SelectionStart + bx.SelectionLength - 1), str2, Right(bx.Text, Len(bx.Text) - bx.SelectionStart - bx.SelectionLength))
                    Else
                        bx.Text = String.Concat(Left(bx.Text, bx.SelectionStart - 1), str2, Right(bx.Text, Len(bx.Text) - bx.SelectionStart))
                    End If
                    bx.SelectionStart = ss
                    bx.SelectionLength = sl
                Else
                    flag = True
                End If
                If flag Then
                    str3 = Mid(bx.Text, bx.SelectionStart + 1, bx.SelectionLength)
                    If str3.Contains(".") Then
                        num2 = Len(str3) - str3.IndexOf(".")
                        i = 2
                        str2 = "."
                        While i <= num2
                            str2 = String.Concat(str2, "0")
                            i += 1
                        End While
                        num2 = str3.IndexOf(".")
                        If Left(str3, 1) = "-" Then num2 -= 1
                        i = 1
                        While i <= num2
                            str2 = String.Concat("0", str2)
                            i += 1
                        End While
                        If CDec(str3) = 0 AndAlso num1 = -1 Then
                            If str3.Contains("-") Then
                                str3 = Left(str3, Len(str3) - 1) + "1"
                            Else
                                str3 = Right(String.Concat(Format(CDec(String.Concat("1", str3)) - Math.Pow(10, 1 + str3.IndexOf(".") - Len(str3)), str2), CStr(IIf(Right(str3, 1) = ".", ".", ""))), Len(str3))
                            End If
                        ElseIf num1 <> -1 Then
                            str3 = If(Right(str3, 1) <> ".", Right(Format(CDec(str3) + Math.Pow(10, 1 + str3.IndexOf(".") - Len(str3)), str2), Len(str3)), Right(String.Concat(Format(CDec(str3) + Math.Pow(10, 1 + str3.IndexOf(".") - Len(str3)), str2), "."), Len(str3)))
                        Else
                            str3 = Right(String.Concat(Format(CDec(str3) - Math.Pow(10, 1 + str3.IndexOf(".") - Len(str3)), str2), CStr(IIf(Right(str3, 1) = ".", ".", ""))), CInt(IIf(str3.Contains("-"), Len(str3) + 1, Len(str3))))
                        End If
                    ElseIf CDec(str3) = 0 AndAlso num1 = -1 Then
                        str3 = CStr(CDec(String.Concat("1", str3)) - 1)
                    ElseIf num1 = -1 Then
                        str2 = ""
                        num2 = Len(str3)
                        If str3.Contains("-") Then num2 -= 1
                        i = 1
                        While i <= num2
                            str2 = String.Concat(str2, "0")
                            i += 1
                        End While
                        str3 = Format(CDec(str3) - 1, str2)
                    ElseIf Len(CStr(CDec(str3) + 1)) <> Len(str3) + 1 Then
                        str2 = ""
                        num2 = Len(str3)
                        If str3.Contains("-") Then num2 -= 1
                        If Math.Log10(-CDec(str3)) Mod 1 = 0 Then num2 -= 1
                        i = 1
                        While i <= num2
                            str2 = String.Concat(str2, "0")
                            i += 1
                        End While
                        str3 = Format(CDec(str3) + 1, str2)
                    Else
                        num2 = Len(str3)
                        str3 = ""
                        i = 1
                        While i <= num2
                            str3 = String.Concat(str3, "0")
                            i += 1
                        End While
                    End If
                    bx.Text = String.Concat(str, str3, str1)
                    bx.SelectionStart = ss
                    bx.SelectionLength = sl
                    Dim k As Boolean
                    If Decimal.TryParse(st, ts) Then
                        If ts < 0 Then
                            If Math.Log10(-CDec(str3)) Mod 1 = 0 AndAlso num1 < 0 Then
                                bx.SelectionLength += 1
                            ElseIf num1 > 0 Then
                                str3 = CStr(-CDec(str3))
                                For j = 1 To Len(str3)
                                    If Mid(str3, j, 1) <> "9" AndAlso Mid(str3, j, 1) <> "." Then
                                        k = True
                                        Exit For
                                    End If
                                Next
                                If Not k OrElse CDec(str3) = 0 Then
                                    If bx.SelectionLength = sl Then bx.SelectionLength -= 1
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                bx.Text = "0"
                bx.SelectionStart = 1
            End If
        Catch ex As Exception
            bx.Text = tt
        End Try
    End Sub
    Sub s51(ByRef T As TextBox, ByRef e As KeyEventArgs, ByRef cnctk As SqlConnection, ByRef bl As Boolean)
        If e.KeyCode = Keys.Enter Then
            Try
                cnctk.Open()
                cmdstr = "select " & T.Text
                cmd = New SqlCommand(cmdstr, cnctk)
                drm = cmd.ExecuteReader
                While drm.Read
                    T.Text = CStr(drm(0))
                End While
                cnctk.Close()
            Catch ex As Exception
                cnctk.Close()
                If bl Then T.Text = ""
            End Try
            T.SelectionStart = T.Text.Length
        ElseIf e.KeyCode = Keys.Escape Then
            T.Text = CStr(T.Tag)
            T.SelectionStart = T.Text.Length
        Else
            T.Tag = T.Text
        End If
    End Sub
    Sub s25(ByRef DA As DataGridView, ByRef cmdstrn As String, ByRef table As String)
        Try
            DA.Columns.Clear()
            cnct.Open()
            drm = New SqlCommand("select 操作员,操作时间,SQL语句,记录表 as 表,记录Id as Id,Id as RId,计算机名 from 操作记录 where 记录Id=" & cmdstrn & " and 记录表 ='" & table & "'", cnct).ExecuteReader
            For i = 0 To drm.FieldCount - 1
                DA.Columns.Add(drm.GetName(i), drm.GetName(i))
            Next
            RemoveHandler Form1.DA11.CellValueChanged, AddressOf Form1.DA11_CellValueChanged
            While drm.Read
                DA.Rows.Add()
                For i = 0 To drm.FieldCount - 1
                    DA.Rows(DA.Rows.Count - 2).Cells(i).Value = s19(CStr(drm(i)), table)
                Next
            End While
            AddHandler Form1.DA11.CellValueChanged, AddressOf Form1.DA11_CellValueChanged
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
    Function s19(ByRef str As String, ByRef table As String) As String
        s19 = Replace(str, "insert into " & table & " values", "",, 1)
        s19 = Replace(s19, "'''", "''")
        s19 = Replace(s19, "''", vbCrLf)
        s19 = Replace(s19, "'", "")
        s19 = Replace(s19, vbCrLf, "'")
    End Function
End Module
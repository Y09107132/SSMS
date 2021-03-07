Imports System.Data.SqlClient
Public Class Form0
    Public js As String
    Dim cmdstr As String
    Dim cmd As SqlCommand
    Dim dr As SqlDataReader
    Public blf, gx As Boolean
    Public st() As String = My.Settings.setting.Split(CChar("|"))
    Dim cnct As SqlConnection = New SqlConnection("data source=" & st(3) & ";initial catalog=" & st(2) & ";user id=" & st(0) & ";password=" & st(1))
    Private Sub Form0_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            cnct.Open()
            cmd = New SqlCommand("select name from sys.sql_logins where principal_id=1", cnct)
            dr = cmd.ExecuteReader
            While dr.Read
                C2.Items.Add(dr(0))
            End While
            cnct.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
            Application.Exit()
        End Try
        cmdstr = "sp_helpuser"
        Try
            cnct.Open()
            cmd = New SqlCommand(cmdstr, cnct)
            dr = cmd.ExecuteReader
            While dr.Read
                If CInt(dr(5)) > 4 Then C2.Items.Add(dr(0))
            End While
            cnct.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
            Application.Exit()
        End Try
        cmdstr = "select name from sys.sql_logins where is_disabled=0 and name<>'calc'"
        Try
            cnct.Open()
            cmd = New SqlCommand(cmdstr, cnct)
            dr = cmd.ExecuteReader
            While dr.Read
                C1.Items.Add(dr(0))
            End While
            cnct.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
            Application.Exit()
        End Try
        Try
            cnct.Open()
            cmd = New SqlCommand("sp_helprole", cnct)
            dr = cmd.ExecuteReader
            While dr.Read
                If CInt(dr(1)) <= 16384 AndAlso CInt(dr(1)) > 0 Then C3.Items.Add(dr(0))
            End While
            cnct.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
            Application.Exit()
        End Try
        C3.Items.Add("")
        Try
            cnct.Open()
            dr = New SqlCommand("select 考试科目 from 考试科目 where 可用性=1", cnct).ExecuteReader
            While dr.Read
                CL1.Items.Add(dr(0))
            End While
            cnct.Close()
        Catch ex As Exception
            cnct.Close()
        End Try
        s2("注意事项", T1)
        s2("系统标题", L1)
        C1.Items.Remove(st(0))
        If Screen.PrimaryScreen.Bounds.Width <= Width OrElse Screen.PrimaryScreen.Bounds.Height <= Height Then MsgBox("屏幕分辨率不得小于690×436！")
        TT1.SetToolTip(C3, "DataReader有成绩数据访问权限" & vbCrLf & "Joperator有录入数据权限" & vbCrLf & "Soperator有数据访问和录入权限" & vbCrLf & "db_owner是数据库管理员" & vbCrLf & "所有权限归" & st(0) & "拥有")
    End Sub
    Private Sub C2_GotFocus(sender As Object, e As EventArgs) Handles C2.GotFocus
        AcceptButton = B2
    End Sub
    Private Sub C2_TextChanged(sender As Object, e As EventArgs) Handles C2.TextChanged
        Dim cmd As SqlCommand
        Dim dr As SqlDataReader
        RemoveHandler T4.TextChanged, AddressOf L1_Text
        T4.Text = ""
        For j = 1 To CL1.Items.Count
            CL1.SetItemChecked(j - 1, False)
        Next
        cmdstr = "select 密码 from 人员设置 where 操作人员=@操作人员 and 所用电脑=@所用电脑 and 用户名=@用户名"
        Try
            cmd = New SqlCommand(cmdstr, cnct)
            cmd.Parameters.Add(New SqlParameter("@操作人员", C2.Text))
            cmd.Parameters.Add(New SqlParameter("@所用电脑", Environment.MachineName))
            cmd.Parameters.Add(New SqlParameter("@用户名", Environment.UserName))
            cnct.Open()
            dr = cmd.ExecuteReader
            CH1.Checked = False
            While dr.Read
                CH1.Checked = True
                T4.Text = CStr(dr(0))
            End While
            cnct.Close()
        Catch ex As Exception
            cnct.Close()
        End Try
        AddHandler T4.TextChanged, AddressOf L1_Text
        If Fcsb.s1(C2.Text) = 0 OrElse Fcsb.s1(C2.Text) = 1 Then
            CL1.Enabled = True
        Else
            CL1.Enabled = False
        End If
        Try
            cnct.Open()
            cmd = New SqlCommand("select 考试科目 from 用户科目 where 操作人员=@操作人员", cnct)
            cmd.Parameters.Add(New SqlParameter("@操作人员", C2.Text))
            dr = cmd.ExecuteReader
            While dr.Read
                For j = 1 To CL1.Items.Count
                    If CL1.Items(j - 1).ToString = dr(0).ToString Then CL1.SetItemChecked((j - 1), True) : Exit For
                Next
            End While
            cnct.Close()
        Catch ex As Exception
            cnct.Close()
        End Try
    End Sub
    Private Sub B2_Click(sender As Object, e As EventArgs) Handles B2.Click
        If C2.Text.Contains("]") OrElse C2.Text = "calc" OrElse C2.Text = "" OrElse C2.Text.Contains("'") Then MsgBox("用户名格式不规范！") : Exit Sub
        If T4.Text.Contains("'") Then MsgBox("密码中不能包含'") : Exit Sub
        Dim cnctn As New SqlConnection("data source=" & st(3) & ";initial catalog=" & st(2) & ";user id=" & C2.Text & ";password=" & T4.Text)
        Try
            cnctn.Open()
            Dim cmdn As New SqlCommand("select 1", cnctn)
            dr = cmdn.ExecuteReader
            Try
                cmdstr = "delete from 人员设置 where 操作人员=@操作人员 and 所用电脑=@所用电脑 and 用户名=@用户名"
                If CH1.Checked Then
                    cmdstr += vbCrLf & "insert into 人员设置 values(@操作人员,@密码,@所用电脑,@用户名)"
                End If
                cnct.Open()
                cmd = New SqlCommand(cmdstr, cnct)
                cmd.Parameters.Add(New SqlParameter("@操作人员", C2.Text))
                cmd.Parameters.Add(New SqlParameter("@密码", T4.Text))
                cmd.Parameters.Add(New SqlParameter("@所用电脑", Environment.MachineName))
                cmd.Parameters.Add(New SqlParameter("@用户名", Environment.UserName))
                cmd.ExecuteNonQuery()
                cnct.Close()
            Catch ex As Exception
                cnct.Close()
            End Try
            cnctn.Close()
            Form1.Show()
            Close()
        Catch ex As Exception
            cnctn.Close()
            MsgBox("用户 " & C2.Text & " 登录失败")
            T4.Focus()
            Exit Sub
        End Try
        Try
            cnct.Open()
            If Not CBool(New SqlCommand("select -1 from sys.sql_logins where name='calc'", cnct).ExecuteScalar) Then Dim i As Integer = New SqlCommand("exec sp_addlogin 'calc','','msdb' use msdb grant connect to guest", cnct).ExecuteNonQuery
            cnct.Close()
        Catch ex As Exception
            MsgBox("创建计算用账户失败！" & vbCrLf & ex.Message)
            Application.Exit()
        End Try
        End Sub
    Private Sub B1_Click(sender As Object, e As EventArgs) Handles B1.Click
        Dim sign As Boolean = True
        s1(sign)
        If sign Then
            s4(cnct)
        End If
    End Sub
    Private Sub L10_Click(sender As Object, e As EventArgs) Handles L10.Click
        Close()
    End Sub
    Private Sub T1_GotFocus(sender As Object, e As EventArgs) Handles T1.GotFocus
        s2("注意事项", T1)
        AcceptButton = Nothing
    End Sub
    Private Sub T1_LostFocus(sender As Object, e As EventArgs) Handles T1.LostFocus
        s3("注意事项", "公告栏", T1)
    End Sub
    Private Sub L1_Text(sender As Object, e As EventArgs) Handles L1.GotFocus, C2.TextChanged, T4.TextChanged
        If T4.Text = st(1) AndAlso C2.Text = st(0) Then
            If sender Is L1 Then
                s2("系统标题", L1)
                AcceptButton = Nothing
            End If
            L1.ReadOnly = False
        Else
            L1.ReadOnly = True
        End If
        If sender Is L1 Then L1.Tag = L1.Text
    End Sub
    Private Sub L1_LostFocus(sender As Object, e As EventArgs) Handles L1.LostFocus
        s3("系统标题", "标题栏", L1)
    End Sub
    Private Sub C1_GotFocus(sender As Object, e As EventArgs) Handles C1.GotFocus
        AcceptButton = B1
    End Sub
    Public Sub C1_TextChanged(sender As Object, e As EventArgs) Handles C1.TextChanged
        Dim cmd As SqlCommand
        Dim dr As SqlDataReader
        Dim bl As Boolean = True
        T2.Text = "" : T3.Text = "" : T5.Text = ""
        If C1.Items.Contains(C1.Text) Then
            T2.Enabled = True
            cnct.Open()
            cmd = New SqlCommand("exec sp_helprolemember", cnct)
            dr = cmd.ExecuteReader
            While dr.Read
                If CStr(dr(1)) = C1.Text Then
                    C3.Text = CStr(dr(0))
                    bl = False
                    js = C3.Text
                    Exit While
                End If
            End While
            cnct.Close()
        Else
            T2.Enabled = False
        End If
        If bl Then C3.Text = "" : js = ""
        For j = 1 To CL1.Items.Count
            CL1.SetItemChecked((j - 1), False)
        Next
        Try
            cnct.Open()
            cmd = New SqlCommand("select 考试科目 from 用户科目 where 操作人员=@操作人员", cnct)
            cmd.Parameters.Add(New SqlParameter("@操作人员", C1.Text))
            dr = cmd.ExecuteReader
            While dr.Read
                For j = 1 To CL1.Items.Count
                    If CL1.Items(j - 1).ToString = dr(0).ToString Then CL1.SetItemChecked((j - 1), True) : Exit For
                Next
            End While
            cnct.Close()
        Catch ex As Exception
            cnct.Close()
        End Try
    End Sub
    Private Sub CL1_ItemCheck(sender As Object, e As ItemCheckEventArgs) Handles CL1.ItemCheck
        If e.Index = 0 Then
            RemoveHandler CL1.ItemCheck, AddressOf CL1_ItemCheck
            If e.NewValue = CheckState.Checked Then
                For i = 1 To CL1.Items.Count - 1
                    CL1.SetItemChecked(i, True)
                Next
            Else
                For i = 1 To CL1.Items.Count - 1
                    CL1.SetItemChecked(i, False)
                Next
            End If
            AddHandler CL1.ItemCheck, AddressOf CL1_ItemCheck
        Else
            RemoveHandler CL1.ItemCheck, AddressOf CL1_ItemCheck
            s12(DirectCast(sender, CheckedListBox), e)
            AddHandler CL1.ItemCheck, AddressOf CL1_ItemCheck
        End If
    End Sub
    Private Sub B3_Click(sender As Object, e As EventArgs) Handles B3.Click
        If C1.Text.Contains("]") OrElse C1.Text = "calc" OrElse C1.Text = "" OrElse C1.Text.Contains("'") Then MsgBox("登录名格式不规范！") : Exit Sub
        If T2.Text.Contains("'") OrElse T3.Text.Contains("'") OrElse T5.Text.Contains("'") Then MsgBox("密码格式不规范！") : Exit Sub
        If C2.Text.Contains("]") OrElse C2.Text = "calc" OrElse C2.Text = "" OrElse C2.Text.Contains("'") Then MsgBox("登录名格式不规范！(管理员)") : Exit Sub
        If T4.Text.Contains("'") Then MsgBox("密码格式不规范！(管理员)") : Exit Sub
        Dim cnct As New SqlConnection("data source=" & st(3) & ";initial catalog=" & st(2) & ";user id=" & C2.Text & ";password=" & T4.Text)
        Dim cmd As SqlCommand
        If MsgBox("是否删除用户？", MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
            Try
                cnct.Open()
                cmd = New SqlCommand("drop user [" & C1.Text & "]", cnct)
                cmd.ExecuteNonQuery()
                cnct.Close()
                MsgBox("删除用户成功！")
                C2.Items.Remove(C1.Text)
                T2.Text = ""
                T3.Text = ""
                T5.Text = ""
                C3.Text = ""
                js = ""
            Catch ex As Exception
                cnct.Close()
                MsgBox("删除用户失败！" & Replace(ex.Message, "'", ""))
            End Try
        End If
        If MsgBox("是否删除登录名？", MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
            Try
                cnct.Open()
                cmd = New SqlCommand("drop login [" & C1.Text & "]", cnct)
                cmd.ExecuteNonQuery()
                cnct.Close()
                MsgBox("删除登录名成功！")
                C1.Items.Remove(C1.Text)
            Catch ex As Exception
                cnct.Close()
                MsgBox("删除登录名失败！" & Replace(ex.Message, "'", ""))
            End Try
        End If
        End Sub
    Sub s1(ByRef sign As Boolean)
        If C1.Text.Contains("]") OrElse C1.Text = "calc" OrElse C1.Text = "" OrElse C1.Text.Contains("'") Then MsgBox("登录名格式不规范！") : sign = False : Exit Sub
        If T2.Text.Contains("'") OrElse T3.Text.Contains("'") OrElse T5.Text.Contains("'") Then MsgBox("密码格式不规范！") : sign = False : Exit Sub
        If C2.Text.Contains("]") OrElse C2.Text = "calc" OrElse C2.Text = "" OrElse C2.Text.Contains("'") Then MsgBox("登录名格式不规范！(管理员)") : sign = False : Exit Sub
        If T4.Text.Contains("'") Then MsgBox("密码格式不规范！(管理员)") : sign = False : Exit Sub
        Dim cnctn As SqlConnection = New SqlConnection("data source=" & st(3) & ";initial catalog=" & st(2) & ";user id=" & C2.Text & ";password=" & T4.Text)
        Dim bl As Boolean = False
        If T2.Enabled Then
            If T2.Text <> "" OrElse T3.Text <> "" OrElse T5.Text <> "" Then
                If T3.Text = T5.Text Then
                    Try
                        cmd = New SqlCommand("sp_password", cnct)
                        cmd.CommandType = CommandType.StoredProcedure
                        cmd.Parameters.Add(New SqlParameter("@old", T2.Text))
                        cmd.Parameters.Add(New SqlParameter("@new", T3.Text))
                        cmd.Parameters.Add(New SqlParameter("@loginame", C1.Text))
                        cnct.Open()
                        cmd.ExecuteNonQuery()
                        MsgBox("密码更改成功！")
                        T2.Text = "" : T3.Text = "" : T5.Text = ""
                    Catch ex As Exception
                        MsgBox("更改密码时发生错误，旧密码错误或者连接失败！")
                        cnct.Close()
                        sign = False
                        Exit Sub
                    End Try
                    cnct.Close()
                Else
                    MsgBox("两次密码不一致，请重输")
                    sign = False
                    Exit Sub
                End If
            End If
        Else
            If T3.Text <> T5.Text Then MsgBox("两次密码输入不一致！") : sign = False : Exit Sub
            Try
                cnctn.Open()
                cmd = New SqlCommand("sp_addlogin", cnctn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add(New SqlParameter("@loginame", C1.Text))
                cmd.Parameters.Add(New SqlParameter("@passwd", T3.Text))
                cmd.ExecuteNonQuery()
                T2.Enabled = True
                If Not C1.Items.Contains(C1.Text) Then C1.Items.Add(C1.Text)
                cnctn.Close()
                MsgBox("成功建立登录名！")
                T3.Text = "" : T5.Text = ""
            Catch ex As Exception
                cnctn.Close()
                MsgBox("未能建立登录名！" & vbCrLf & Replace(ex.Message, "'", ""))
                sign = False
                Exit Sub
            End Try
            If C3.Text = "" Then js = "" : C2.Items.Remove(C1.Text) : sign = False : Exit Sub
        End If
        RemoveHandler C1.TextChanged, AddressOf C1_TextChanged
        If Fcsb.s1(C1.Text) = 1 AndAlso Fcsb.s1(C2.Text) = 1 Then
            MsgBox("角色db_owner不能相互更改。")
            sign = False
            C3.Text = js
            AddHandler C1.TextChanged, AddressOf C1_TextChanged
            Exit Sub
        End If
        AddHandler C1.TextChanged, AddressOf C1_TextChanged
        If Fcsb.s1(C2.Text) = 1 AndAlso C3.Text = "db_owner" Then
            If Global.Microsoft.VisualBasic.Interaction.MsgBox("你正在尝试将 " & Me.C1.Text & " 的权限升高到报表数据库的最高级别" & Global.Microsoft.VisualBasic.Constants.vbCrLf & "一旦生效你将没有权限将其降级，是否继续？", Global.Microsoft.VisualBasic.MsgBoxStyle.YesNo) = Global.Microsoft.VisualBasic.MsgBoxResult.No Then
                bl = True
            End If
        End If
        If bl Then
            sign = False
            C3.Text = js
            Exit Sub
        End If
        If js = C3.Text Then
            If js = "" Then
                sign = False
            End If
            Exit Sub
        End If
        Try
            cnctn.Open()
            cmd = New SqlCommand("drop user [" & C1.Text & "]", cnctn)
            cmd.ExecuteNonQuery()
            cnctn.Close()
            If C3.Text = "" AndAlso js <> C3.Text Then
                MsgBox("去除角色成功！")
                js = ""
                C2.Items.Remove(C1.Text)
                C3.Text = ""
                sign = False
                Exit Sub
            End If
        Catch ex As Exception
            cnctn.Close()
        End Try
        Try
            cnctn.Open()
            cmd = New SqlCommand("create user [" & C1.Text & "] for login [" & C1.Text & "] with default_schema=dbo", cnctn)
            cmd.ExecuteNonQuery()
            cnctn.Close()
        Catch ex As Exception
            cnctn.Close()
        End Try
        Try
            cnctn.Open()
            cmd = New SqlCommand("sp_addrolemember", cnctn)
            cmd.Parameters.Add(New SqlParameter("@rolename", C3.Text))
            cmd.Parameters.Add(New SqlParameter("@membername", C1.Text))
            cmd.CommandType = CommandType.StoredProcedure
            cmd.ExecuteNonQuery()
            cnctn.Close()
            MsgBox("角色赋予成功！")
            If js = "" Then
                gx = True
            End If
            If Not C2.Items.Contains(C1.Text) Then C2.Items.Add(C1.Text)
            js = C3.Text
        Catch ex As Exception
            cnctn.Close()
            MsgBox("角色未能正常更改！" & vbCrLf & Replace(ex.Message, "'", ""))
            C3.Text = js
            sign = False
            Exit Sub
        End Try
        End Sub
    Sub s2(ByRef str As String, ByRef T As TextBox)
        cmdstr = "select " & str & " from 系统配置"
        Try
            cnct.Open()
            cmd = New SqlCommand(cmdstr, cnct)
            dr = cmd.ExecuteReader
            While dr.Read
                T.Text = CStr(dr(0))
            End While
            cnct.Close()
        Catch ex As Exception
            cnct.Close()
            MsgBox(ex.Message)
            Application.Exit()
        End Try
    End Sub
    Sub s3(ByRef str1 As String, ByRef str2 As String, ByRef T As TextBox)
        cmdstr = "update 系统配置 set " & str1 & "=@Content"
        Try
            cnct.Open()
            cmd = New SqlCommand(cmdstr, cnct)
            cmd.Parameters.Add(New SqlParameter("@Content", T.Text))
            cmd.ExecuteNonQuery()
            cnct.Close()
        Catch ex As Exception
            cnct.Close()
            MsgBox(str2 & "可能没有正确更新" & vbCrLf & ex.Message)
            Exit Sub
        End Try
    End Sub
    Private Sub L1_KeyDown(sender As Object, e As KeyEventArgs) Handles L1.KeyDown
        If e.KeyCode = Keys.Escape Then DirectCast(sender, TextBox).Text = CStr(DirectCast(sender, TextBox).Tag)
    End Sub
    Sub s4(ByRef cnctm As SqlConnection)
        Dim cnctn As SqlConnection = New SqlConnection("data source=" & st(3) & ";initial catalog=" & st(2) & ";user id=" & C2.Text & ";password=" & T4.Text)
        Dim da As SqlDataAdapter
        Dim dt1 As New DataTable
        Dim dt2 As New DataTable
        Dim bl As Boolean
        Dim bn As Boolean
        cmdstr = "select 考试科目 from 用户科目 where 操作人员=@操作人员 order by 考试科目"
        cmd = New SqlCommand(cmdstr, cnctm)
        cmd.Parameters.Add(New SqlParameter("@操作人员", C1.Text))
        da = New SqlDataAdapter(cmd)
        da.Fill(dt1)
        Try
            cnctn.Open()
            cmdstr = "delete from 用户科目 where 操作人员=@操作人员"
            cmd = New SqlCommand(cmdstr, cnctn)
            cmd.Parameters.Add(New SqlParameter("@操作人员", C1.Text))
            cmd.ExecuteNonQuery()
            cnctn.Close()
            bn = True
        Catch ex As Exception
            cnctn.Close()
        End Try
        dt2.Columns.Add("考试科目")
        For i = 1 To CL1.CheckedItems.Count - 1
            dt2.Rows.Add(CL1.CheckedItems(i).ToString)
            cmdstr = "insert into 用户科目 Values(@操作人员,@考试科目)"
            Try
                cnctn.Open()
                cmd = New SqlCommand(cmdstr, cnctn)
                cmd.Parameters.Add(New SqlParameter("@操作人员", C1.Text))
                cmd.Parameters.Add(New SqlParameter("@考试科目", CL1.CheckedItems(i).ToString))
                cmd.ExecuteNonQuery()
                cnctn.Close()
            Catch ex As Exception
                cnctn.Close()
            End Try
        Next
        Dim dtr() As DataRow = dt2.Select("", "考试科目 asc")
        If gx Then
            bl = True
            gx = False
        ElseIf dt1.Rows.Count = dt2.Rows.Count Then
            For i = 0 To dt1.Rows.Count - 1
                If dtr(i)(0).ToString <> dt1.Rows.Item(i)(0).ToString Then
                    bl = True
                    Exit For
                End If
            Next
        Else
            bl = True
        End If
        If bl Then
            If bn Then
                MsgBox("用户科目更新成功！")
            Else
                MsgBox("没有足够的权限更新用户科目！")
            End If
        End If
    End Sub
End Class
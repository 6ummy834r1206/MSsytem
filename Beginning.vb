Imports System.Data.SqlClient
Public Class Beginning
    Public cnStr As String = "Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\DBNAME;Integrated Security=True;Connect Timeout=30"
    Public reason As String = "換班"
    Public changeWork As Integer = 0
    Private Sub Beginning_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Focus()
        ComboBox1.DropDownStyle = ComboBoxStyle.DropDownList
        TextBox2.PasswordChar = "*"
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox1.Focus()
    End Sub

    Private Sub Reason_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        reason = ComboBox1.Text
    End Sub

    Private Sub Enter_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Dim cn As New SqlConnection(cnStr)
            cn.Open()
            'Dim commandtext As String = "SELECT * From MSuser Where account ='" + TextBox1.Text + "' AND password = '" + TextBox2.Text + "'"
            Dim commandtext As String = "SELECT * From MSuser Where account =@Account AND password = @Password"
            Using connection As New SqlConnection(cnStr)
                Using command As New SqlCommand(commandtext, connection)
                    Try
                        ' 添加参数
                        command.Parameters.AddWithValue("@Account", TextBox1.Text)
                        command.Parameters.AddWithValue("@Password", TextBox2.Text)
                        ' 打开连接
                        connection.Open()
                        ' 执行查询
                        Using reader As SqlDataReader = command.ExecuteReader()
                            ' 处理查询结果
                            If reader.HasRows Then
                                If Main.pauseClick Then
                                    If ComboBox1.Text = "" Then
                                        MsgBox("請輸入停機原因")
                                    Else
                                        Main.Show()
                                        Main.WritePause()
                                        If ComboBox1.Text.Equals("換班") Then
                                            changeWork = 1
                                            Main.Button0.Enabled = True
                                            Work.ShowDialog()
                                        End If
                                        Me.Hide()
                                        TextBox2.Text = ""
                                    End If
                                Else
                                    Me.Hide()
                                    Main.Show()
                                    TextBox2.Text = ""
                                End If
                            Else
                                MsgBox("帳號密碼錯誤")
                                TextBox2.Text = ""
                                cn.Close()
                            End If
                        End Using
                    Catch ex As Exception
                        ' 处理异常
                        MessageBox.Show("Error: " & ex.Message)
                    End Try
                End Using
            End Using
            '    Dim cmd As New SqlCommand(commandtext, cn)
            '    Dim dr As SqlDataReader
            '    dr = cmd.ExecuteReader
            '    If dr.HasRows Then
            '        If Main.pauseClick Then
            '            If ComboBox1.Text = "" Then
            '                MsgBox("請輸入停機原因")
            '            Else
            '                Main.Show()
            '                Main.WritePause()
            '                If ComboBox1.Text.Equals("換班") Then
            '                    changeWork = 1
            '                    Main.Button0.Enabled = True
            '                    Work.ShowDialog()
            '                End If
            '                Me.Hide()
            '                TextBox2.Text = ""
            '            End If
            '        Else
            '            Me.Hide()
            '            Main.Show()
            '            TextBox2.Text = ""
            '        End If
            '    Else
            '        MsgBox("帳號密碼錯誤")
            '        TextBox2.Text = ""
            '    End If
            '    cn.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If e.KeyChar = ChrW(Keys.Return) Then
            Enter_Click(sender, e)
        End If
    End Sub

    Private Sub Beginning_Activated(sender As Object, e As EventArgs) Handles MyBase.Activated
        If Main.pauseClick Then
            Main.Hide()
            Label3.Visible = True
            ComboBox1.Visible = True
            Button2.Visible = False
        Else
            Label3.Visible = False
            ComboBox1.Visible = False
            Button2.Visible = True
        End If
        If ComboBox1.Items.Count = 0 Then
            SetReason()
        End If
    End Sub

    Private Sub Beginning_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If Main.pauseClick Then
            Main.WritePause()
        End If
        Application.Exit()
    End Sub

    Private Sub Leave_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Application.Exit()
    End Sub

    Private Sub SetReason()
        Dim cn As New SqlConnection(Main.cnStr)
        Try
            cn.Open()
            Dim commandtext As String = "SELECT * From PauseReason"
            Dim cmd As New SqlCommand(commandtext, cn)
            Dim dr As SqlDataReader
            dr = cmd.ExecuteReader
            While (dr.Read())
                ComboBox1.Items.Add(Trim(dr("Reason").ToString))
            End While
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cn.Close()
        End Try
    End Sub
End Class
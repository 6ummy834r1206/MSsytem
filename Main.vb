Imports System.Data.SqlClient
Imports System.Net.Sockets
Imports System.IO
Imports System.Text
Imports Microsoft.Office.Interop
Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Windows.Forms
Imports System.Threading

Public Class Main

    Public cnStr As String = "Data Source=;AttachDbFilename=;Integrated Security=True;Connect Timeout=30"
    Public tcp As New TcpClient
    Public heightN As Integer = 1
    Public IDD As Integer = 1
    Public heightOK As Integer = 0
    Public heightNG As Integer = 0
    Public h As String = ""
    Public ccdNG As Integer = 0
    Public pauseTime As String = ""
    Dim usl As Double = "0"
    Dim lsl As Double = "0"
    Public pauseClick As Integer = 0
    Public heightInfo As String = ""
    Private Delegate Sub dlgObjs()
    Public Declare Function GetTickCount Lib "kernel32" () As Long
    Public ht As Integer = 3
    Public htOK As Integer = 3
    Public htNG As Integer = 3
    Dim newWork As Integer = 0
    Public stopReason As String = ""
    Public errcode As Integer = 0
    Dim ms As Integer = 0
    Private Sub Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Timer1.Start()
        DataGridView1.RowHeadersVisible = False
        NewMonth()
        ButtonInitial()
        LastWork()
        GetUslLsl()
        Label19.ForeColor = Color.Black
        Label19.Text = "請新增工單"
        Me.UseWaitCursor = False
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub ButtonInitial()
        Button0.Enabled = True
        Button1.Enabled = False
        Button2.Enabled = False
        Button3.Enabled = True
        Button4.Enabled = False
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Timer1.Interval = 1000
    End Sub

    Public Sub TPNewWork()
        Try
            If tcp.Connected = True Then
                Dim sendbuf As Byte() = New Byte() {}
                tcp.GetStream.Write(sendbuf, 0, sendbuf.Length)
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub TPEndWork()
        Try
            If tcp.Connected = True Then
                Dim sendbuf As Byte() = New Byte() {}
                tcp.GetStream.Write(sendbuf, 0, sendbuf.Length)
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub TPRunning()
        Try
            If tcp.Connected = True Then
                Dim sendbuf As Byte() = New Byte() {}
                tcp.GetStream.Write(sendbuf, 0, sendbuf.Length)
            Else
                MsgBox("請確認連線是否正常")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub TPStopping()
        Try
            If tcp.Connected = True Then
                Dim sendbuf As Byte() = New Byte() {}
                tcp.GetStream.Write(sendbuf, 0, sendbuf.Length)
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub NewWork_Click(sender As Object, e As EventArgs) Handles Button0.Click
        Work.ShowDialog()
    End Sub

    Private Sub Start_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Button1.Enabled = False
        Button2.Enabled = True
        Button3.Enabled = False
        Button4.Enabled = True
        If pauseClick Then
            WriteRestart()
            pauseClick = 0
        End If
        ErrGone()
        delay(200)
        StopGone()
        BackgroundWorker1.RunWorkerAsync()
    End Sub

    Private Sub Pause_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Label19.Text = "待機中"
        Label19.ForeColor = Color.Black
        Button3.Enabled = True
        Button2.Enabled = False
        Button1.Enabled = True
        Beginning.Button2.Visible = False
        Pause()
    End Sub

    Private Sub Setting_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Password.ShowDialog()
        Me.Focus()
    End Sub

    Private Sub EndWork_Click(sender As Object, e As EventArgs) Handles Button4.Click
        BackgroundWorker1.CancelAsync()
        Label19.ForeColor = Color.Red
        Label19.Text = "資料保存中"
        newWork = 1
        Button2.Enabled = False
        MovePicture()
        CCDToCsv()
        WriteHeightToCsv()
        TPChangeWork()
        delay(200)
        TPStopping()
        delay(200)
        TPEndWork()
        delay(200)
        TPErr()
        delay(200)
        TPStop()
        Work.InitialData()
        Work.workPath = ""
        ButtonInitial()
        Beginning.ComboBox1.Text = ""
        Label19.ForeColor = Color.Black
        Label19.Text = "請新增工單"
    End Sub

    Public Function ConnectOn()
        Try
            If tcp.Connected = False Then
                tcp.Connect("", )
            End If
        Catch ex As SocketException
            MsgBox(ex.Message)
            Return False
        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        End Try
        Try
            If SerialPort1.IsOpen = False Then
                SerialPort1.Open()
            End If
        Catch ex As IOException
            MsgBox("連接失敗", 16, "錯誤")
            Return False
        Catch ex As InvalidOperationException
            MsgBox("連接已開啟")
            Return False
        End Try
        Return True
    End Function

    Public Sub GetUslLsl()
        Dim cn As New SqlConnection(cnStr)
        Try
            cn.Open()
            Dim commandtext As String = "SELECT * From UslLsl"
            Dim cmd As New SqlCommand(commandtext, cn)
            Dim dr As SqlDataReader
            dr = cmd.ExecuteReader
            Dim dt As DataTable = New DataTable
            dt.Load(dr)
            usl = dt.Rows(0).Item(0)
            lsl = dt.Rows(0).Item(1)
            If Not usl.ToString.Equals(Label17.Text) Or Not lsl.ToString.Equals(Label18.Text) Then
                Label17.Text = "+" + usl.ToString
                Label18.Text = "-" + lsl.ToString
                LineChart()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cn.Close()
        End Try
    End Sub

    Private Sub WriteRestart()
        Dim filePath As String = "".xlsx"
        Try
            Dim excelApp As Excel.Application
            Dim excelWorksheet As Excel.Worksheet
            Dim excelBook As Excel.Workbook
            excelApp = CreateObject("Excel.Application")
            excelBook = excelApp.Workbooks.Open(filePath)
            excelApp.Visible = False
            excelWorksheet = excelApp.ActiveWorkbook.Sheets("ErrCode")
            excelWorksheet.Activate()
            Dim restartInfo As String() = {Now().ToString, "'" + Label1.Text, "復機", "", "'" + Label2.Text}
            Dim lastRow As Integer = excelWorksheet.UsedRange.Rows.Count + 1
            For i As Integer = 1 To restartInfo.Length
                excelWorksheet.Cells(lastRow, i) = restartInfo(i - 1)
            Next
            excelWorksheet.Columns.EntireColumn.AutoFit()
            excelApp.DisplayAlerts = False
            excelBook.Save()
            excelBook.Close(True)
            excelApp.Quit()
            excelApp = Nothing
            excelBook = Nothing
            excelWorksheet = Nothing
            GC.Collect()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        CheckWork()
    End Sub

    Private Sub CheckWork()
        If ConnectOn() Then
            delay(200)
            TPRunning()
            Label19.ForeColor = Color.Green
            Label19.Text = "運行中"
            Dim run As Integer = 1
            While True
                Try
                    If tcp.Connected = True Then
                        Dim sendbuf As Byte() = New Byte() {}
                        tcp.GetStream.Write(sendbuf, 0, sendbuf.Length)
                        delay(200)
                        If tcp.GetStream.CanRead = True Then
                            Dim readbuf(tcp.ReceiveBufferSize) As Byte
                            tcp.GetStream.Read(readbuf, 0, tcp.ReceiveBufferSize)
                            run = BitConverter.ToInt32(readbuf, 2)
                            Select Case run
                                Case '急停
                                    StopM()
                                    Label19.ForeColor = Color.Green
                                    Label19.Text = "運行中"
                                    delay(200)
                                Case '故障
                                    Err()
                                    Label19.ForeColor = Color.Green
                                    Label19.Text = "運行中"
                                    delay(200)
                                Case 
                                    Label19.ForeColor = Color.Black
                                    Label19.Text = "待機中"
                                    MsgBox("請先停機，再重新啟動", 16, "錯誤")
                            End Select
                        Else
                            MsgBox("無傳回資料", 16, "錯誤")
                        End If
checkCCD:
                        sendbuf = New Byte() {}
                        tcp.GetStream.Write(sendbuf, 0, sendbuf.Length)
                        delay(200)
                        If tcp.GetStream.CanRead = True Then
                            Dim readbuf(tcp.ReceiveBufferSize) As Byte
                            tcp.GetStream.Read(readbuf, 0, tcp.ReceiveBufferSize)
                            ccdNG = BitConverter.ToInt32(readbuf, 2)
                            If ccdNG > Label4.Text Then
                                GoTo checkCCD
                            End If
                        Else
                            MsgBox("無傳回資料", 16, "錯誤")
                        End If
                    Else
                        MsgBox("連接中斷", 16, "錯誤")
                        Button1.Enabled = True
                        Button2.Enabled = False
                        Label19.ForeColor = Color.Black
                        Label19.Text = "待機中"
                        Exit Sub
                    End If
                Catch ex As Exception
                    MsgBox(ex.Message)
                    tcp.Close()
                    tcp = New TcpClient
                End Try
                Dim origin As String = Convert.ToString(heightOK + ccdNG + heightNG)
                If Not Label4.Text.Equals(origin) And origin < 99999 Then
                    Label5.Text = heightOK
                    Label15.Text = heightNG
                    Label4.Text = origin
                    Label16.Text = ccdNG
                    WriteCCD()
                End If
                If newWork Then
                    newWork = 0
                    Exit Sub
                End If
                If pauseClick Then
                    Exit Sub
                End If
            End While
        Else
            MsgBox("連線異常")
            Button1.Enabled = True
            Button2.Enabled = False
            Label19.ForeColor = Color.Black
            Label19.Text = "待機中"
        End If
    End Sub

    Private Sub WriteCCD()
        Dim cn As New SqlConnection(cnStr)
        Try
            cn.Open()
            Dim commandtext1 As String = ""
            Dim cmd1 As New SqlCommand(commandtext1, cn)
            cmd1.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cn.Close()
        End Try
    End Sub

    Private Sub Pause()
        pauseClick = 1
        TPStopping()
        delay(200)
        TPErr()
        delay(200)
        TPStop()
        pauseTime = Now().ToString
        Beginning.Visible = True
    End Sub

    Public Sub WritePause()
        Dim filePath As String = "" + ".xlsx"
        Try
            Dim excelApp As Excel.Application
            Dim excelWorksheet As Excel.Worksheet
            Dim excelBook As Excel.Workbook
            excelApp = CreateObject("Excel.Application")
            excelBook = excelApp.Workbooks.Open(filePath)
            excelApp.Visible = False
            excelWorksheet = excelApp.ActiveWorkbook.Sheets("ErrCode")
            excelWorksheet.Activate()
            Dim errInfo As String() = {pauseTime, "'" + Label1.Text, Beginning.reason, "", "'" + Label2.Text}
            Dim lastRow As Integer = excelWorksheet.UsedRange.Rows.Count + 1
            For i As Integer = 1 To errInfo.Length
                excelWorksheet.Cells(lastRow, i) = errInfo(i - 1)
            Next
            excelWorksheet.Columns.EntireColumn.AutoFit()
            excelApp.DisplayAlerts = False
            excelBook.Save()
            excelBook.Close(True)
            excelApp.Quit()
            excelApp = Nothing
            excelBook = Nothing
            excelWorksheet = Nothing
            GC.Collect()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub TPStop()
        Try
            If tcp.Connected = True Then
                Dim sendbuf As Byte() = New Byte() {}
                tcp.GetStream.Write(sendbuf, 0, sendbuf.Length)
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub StopM()
        Label19.ForeColor = Color.Red
        Label19.Text = "急停"
        Dim stopTime As String = Now().ToString
        FindErrCode()
        delay(200)
        TPStop()
        SearchErrTable(errcode)
        Dim pwd As New Password
        pwd.page = 1
        pwd.ShowDialog()
        pwd.Close()
        Dim sol1 As New Solution
        sol1.code = errcode
        sol1.reason = stopReason
        sol1.solType = 1
        sol1.ShowDialog()
        WriteStop(stopTime, sol1.solution)
        sol1.Close()
        StopGone()
    End Sub

    Private Sub WriteStop(ByVal stopTime As String, ByVal solution As String)
        Dim filePath As String = "" + ".xlsx"
        Try
            Dim excelApp As Excel.Application
            Dim excelWorksheet As Excel.Worksheet
            Dim excelBook As Excel.Workbook
            excelApp = CreateObject("Excel.Application")
            excelBook = excelApp.Workbooks.Open(filePath)
            excelApp.Visible = False
            excelWorksheet = excelApp.ActiveWorkbook.Sheets("ErrCode")
            excelWorksheet.Activate()
            Dim stopInfo As String() = {stopTime, "'" + Label1.Text, stopReason, solution, "'" + Label2.Text}
            Dim lastRow As Integer = excelWorksheet.UsedRange.Rows.Count + 1
            For i As Integer = 1 To stopInfo.Length
                excelWorksheet.Cells(lastRow, i) = stopInfo(i - 1)
            Next
            excelWorksheet.Columns.EntireColumn.AutoFit()
            excelApp.DisplayAlerts = False
            excelBook.Save()
            excelBook.Close(True)
            excelApp.Quit()
            excelApp = Nothing
            excelBook = Nothing
            excelWorksheet = Nothing
            GC.Collect()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub StopGone()
        Try
            If tcp.Connected = True Then
                Dim sendbuf As Byte() = New Byte() {}
                tcp.GetStream.Write(sendbuf, 0, sendbuf.Length)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub TPErr()
        Try
            If tcp.Connected = True Then
                Dim sendbuf As Byte() = New Byte() {}
                tcp.GetStream.Write(sendbuf, 0, sendbuf.Length)
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub FindErrCode()
        Try
            If tcp.Connected = True Then
                Dim sendbuf As Byte() = New Byte() {}
                tcp.GetStream.Write(sendbuf, 0, sendbuf.Length)
                delay(200)
                If tcp.GetStream.CanRead = True Then
                    Dim readbuf(tcp.ReceiveBufferSize) As Byte
                    tcp.GetStream.Read(readbuf, 0, tcp.ReceiveBufferSize)
                    errcode = BitConverter.ToInt32(readbuf, 2)
                    If errcode > 74 Then
                        delay(200)
                        FindErrCode()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Sub SearchErrTable(ByVal errcode As Integer)
        Try
            Dim cn As New SqlConnection(cnStr)
            cn.Open()
            Dim commandtext As String = "SELECT * From ErrTable Where ErrCode ="
            Dim cmd As New SqlCommand(commandtext, cn)
            Dim dr As SqlDataReader
            dr = cmd.ExecuteReader
            If dr.HasRows Then
                If (dr.Read()) Then
                    stopReason = dr("ErrReason").ToString
                End If
            Else
                MsgBox("無此故障代碼")
                MsgBox("代碼 : " + Convert.ToString(errcode))
                stopReason = "無此故障代碼"
            End If
            cn.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub WriteErrCode()
        Dim sol As New Solution
        sol.code = errcode
        sol.reason = stopReason
        sol.solType = 0
        sol.ShowDialog()
        Dim filePath As String = ""
        Try
            Dim excelApp As Excel.Application
            Dim excelWorksheet As Excel.Worksheet
            Dim excelBook As Excel.Workbook
            excelApp = CreateObject("Excel.Application")
            excelBook = excelApp.Workbooks.Open(filePath) 
            excelApp.Visible = False
            excelWorksheet = excelApp.ActiveWorkbook.Sheets("ErrCode")
            excelWorksheet.Activate()
            Dim errInfo As String() = {Now().ToString, "'" + Label1.Text, stopReason, sol.solution, "'" + Label2.Text}
            Dim lastRow As Integer = excelWorksheet.UsedRange.Rows.Count + 1
            For i As Integer = 1 To errInfo.Length
                excelWorksheet.Cells(lastRow, i) = errInfo(i - 1)
            Next
            excelWorksheet.Columns.EntireColumn.AutoFit()
            excelApp.DisplayAlerts = False
            excelBook.Save()
            excelBook.Close(True)
            excelApp.Quit()
            excelApp = Nothing
            excelBook = Nothing
            excelWorksheet = Nothing
            sol.Close()
            GC.Collect()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub ErrGone()
        Try
            If tcp.Connected = True Then
                Dim sendbuf As Byte() = New Byte() {}
                tcp.GetStream.Write(sendbuf, 0, sendbuf.Length)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Sub Err()
        Label19.ForeColor = Color.Red
        Label19.Text = "故障"
        FindErrCode()
        delay(200)
        TPErr()
        SearchErrTable(errcode)
        WriteErrCode()
        ErrGone()
    End Sub

    Private Sub DataReceived(ByVal sender As Object, ByVal e As System.IO.Ports.SerialDataReceivedEventArgs) Handles SerialPort1.DataReceived
        Try
            Dim buff(SerialPort1.BytesToRead - 1) As Byte
            SerialPort1.Read(buff, 0, buff.Length)
            Dim msg As String = ""
            Dim sign As String = ""
            Dim number As String = ""
            sign = Chr(buff(3).ToString)
            For i As Integer = 4 To buff.Length - 1
                msg += Chr(buff(i).ToString)
            Next
            msg = msg.Trim("0")
            If msg.Substring(0, 1) = "." Then
                msg = "0" + msg
            End If
            number = sign + msg
            h = number
            heightInfo = HeightOKNG()
            PrintHeight()
            TPOKNG(heightInfo)
            AddtoChart()
            WriteHeightData()
            heightN = heightN + 1
            IDD = IDD + 1
        Catch ex As ArgumentOutOfRangeException
            MsgBox("請確定測量機器已開啟")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Function HeightOKNG()
        If Asc(h) = 43 Then
            If usl.ToString.CompareTo(h) < 0 Then
                heightNG = heightNG + 1
                Return "NG"
            Else
                heightOK = heightOK + 1
                Return "OK"
            End If
        Else
            If lsl.ToString.CompareTo(h) < 0 Then
                heightNG = heightNG + 1
                Return "NG"
            Else
                heightOK = heightOK + 1
                Return "OK"
            End If
        End If
    End Function

    Private Sub TPOKNG(ByVal info As String)
        Try
            If info = "NG" Then
                If tcp.Connected = True Then
                    Dim sendbuf As Byte() = New Byte() {}
                    tcp.GetStream.Write(sendbuf, 0, sendbuf.Length)
                End If
            Else
                If info = "OK" Then
                    If tcp.Connected = True Then
                        Dim sendbuf As Byte() = New Byte() {}
                        tcp.GetStream.Write(sendbuf, 0, sendbuf.Length)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub WriteHeightData()
        Dim cn As New SqlConnection(cnStr)
        Try
            cn.Open()
            Dim commandtext As String = ""
            Dim commandtext1 As String = ""
            If heightInfo.Equals("OK") Then
                commandtext = "INSERT INTO HeightTable(AID,Height,OKNG,ID,Shift,Operator) VALUES (N'" & heightN.ToString("00000") + "','" + h + "','" + heightInfo + "','" + IDD.ToString("0000") + "','" + Label3.Text + "','" + Label2.Text & "')"
                commandtext1 = "INSERT INTO OKTable(OKID,Height) VALUES ( '" & heightOK.ToString("0000") + "','" + h & "')"
            Else
                commandtext = "INSERT INTO HeightTable(AID,Height,OKNG,ID,Shift,Operator) VALUES (N'" & heightN.ToString("00000") + "','" + h + "','" + heightInfo + "','" + IDD.ToString("0000") + "','" + Label3.Text + "','" + Label2.Text & "')"
                commandtext1 = "INSERT INTO NGTable(NGID,Height) VALUES ( '" & heightNG.ToString("0000") + "','" + h & "')"
            End If
            Dim cmd As New SqlCommand(commandtext, cn)
            Dim cmd1 As New SqlCommand(commandtext1, cn)
            cmd.ExecuteNonQuery()
            cmd1.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cn.Close()
        End Try
    End Sub

    Private Sub PrintHeight()
        Dim dgvr As New DataGridViewRow()
        dgvr.CreateCells(Me.DataGridView1)
        dgvr.Cells(0).Value = heightN
        dgvr.Cells(1).Value = h
        Me.DataGridView1.Rows.Add(dgvr)
        Dim i As Integer = DataGridView1.Rows.Count - 1
        DataGridView1.CurrentCell = DataGridView1.Rows(i).Cells(0)
        If heightInfo = "NG" Then
            DataGridView1.Rows(i).Cells(1).Style.ForeColor = Color.Red
        End If
    End Sub

    Private Sub NewMonth()
        Dim folder As String = "d:\"
        Dim killer As String = ""
        If (Directory.Exists(folder)) Then
            If Not File.Exists("d:\") Then
                Dim load As New Loading
                load.Label1.Text = "檔案壓縮中"
                load.Show()
                My7_ZIP()
                load.loading = 0
                delay(2000)
                load.Close()
                MsgBox("請擷取上個月的檔案！", 64)
                Me.Focus()
            End If
            If Month(Now()) = 2 Then
                killer = ""
                If File.Exists(killer) Then
                    My.Computer.FileSystem.DeleteFile(killer)
                End If
            Else
                killer = ""
                If File.Exists(killer) Then
                    My.Computer.FileSystem.DeleteFile(killer)
                End If
            End If
        Else
            folder = ""
            If (Directory.Exists(folder)) Then
                If Not File.Exists("") Then
                    Dim load As New Loading
                    load.Label1.Text = "檔案壓縮中"
                    load.Show()
                    My7_ZIP()
                    load.loading = 0
                    delay(2000)
                    load.Close()
                    MsgBox("請擷取上個月的檔案！", 64)
                    Me.Focus()
                End If
                killer = "d:\" + (Year(Now()) - 1).ToString + "\11.zip"
                If File.Exists(killer) Then
                    My.Computer.FileSystem.DeleteFile(killer)
                End If
            End If
        End If
    End Sub

    Private Sub My7_ZIP()
        Try
            If Month(Now()).ToString <> "1" Then
                Dim myprocess As New Process
                Dim args = ""
                Dim pgm = ""
                With myprocess.StartInfo
                    .FileName = pgm
                    .Arguments = args
                    .WorkingDirectory = "d:\"
                    .UseShellExecute = False
                    .CreateNoWindow = True
                    .RedirectStandardOutput = True
                End With
                myprocess.Start()
                myprocess.BeginOutputReadLine()
                myprocess.WaitForExit()
                myprocess.Close()
            Else
                Dim myprocess As New Process
                Dim args = ""
                Dim pgm = ""
                With myprocess.StartInfo
                    .FileName = pgm
                    .Arguments = args
                    .WorkingDirectory = ""
                    .UseShellExecute = False
                    .CreateNoWindow = True
                    .RedirectStandardOutput = True
                End With
                myprocess.Start()
                myprocess.BeginOutputReadLine()
                myprocess.WaitForExit()
                myprocess.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Sub WriteNewHeight()
        Dim cn As New SqlConnection(cnStr)
        Try
            cn.Open()
            Dim commandtext As String = ""
            commandtext = "INSERT INTO HeightTable(AID,Height,OKNG,ID,Shift,Operator) VALUES(N'" + Now.ToString + "','0','0',N'" + IDD.ToString("0000") + "',N'" + Label3.Text + "','" + Label2.Text + "')"
            Dim cmd As New SqlCommand(commandtext, cn)
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cn.Close()
        End Try
        IDD = IDD + 1
    End Sub

    Private Sub Main_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        SerialPort1.Close()
        tcp.Close()
        Application.Exit()
    End Sub

    Private Sub TPChangeWork()
        Try
            If tcp.Connected = True Then
                Dim sendbuf As Byte() = New Byte() {}
                tcp.GetStream.Write(sendbuf, 0, sendbuf.Length)
            End If
        Catch ex As SocketException
            MsgBox(ex.Message)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub LineChart()
        Chart1.Series.Clear()
        Dim newSeries As New Series("Height")
        With newSeries
            .LegendText = "高度"
            .Color = Color.Red
            .BorderWidth = 2
            .ChartType = SeriesChartType.Line
        End With
        Chart1.ChartAreas(0).AxisY.Interval = 1
        Chart1.ChartAreas(0).AxisX.Maximum = 40
        Chart1.ChartAreas(0).AxisX.MajorGrid.LineWidth = 0
        Chart1.ChartAreas(0).AxisY.LabelStyle.Format = "{0:0.0000}"
        Chart1.ChartAreas(0).AxisX.Interval = 1
        Chart1.Series.Add(newSeries)
        Chart1.ChartAreas(0).AxisY.MajorGrid.Enabled = False
        Chart1.ChartAreas(0).AxisY.MajorTickMark.Enabled = False
        Chart1.ChartAreas(0).AxisY.CustomLabels.Clear()
        Chart1.ChartAreas(0).AxisY.CustomLabels.Add(lsl * 2, 0, "Lsl：" + vbCrLf + lsl.ToString)
        Chart1.ChartAreas(0).AxisY.CustomLabels.Add(0, usl * 2, "Usl：" + vbCrLf + usl.ToString)
        Chart1.ChartAreas(0).AxisY.StripLines.Clear()
        Dim sl1 As StripLine = New StripLine()
        sl1.BackColor = Color.Black
        sl1.IntervalOffset = Label17.Text
        sl1.StripWidth = 0.05
        Chart1.ChartAreas(0).AxisY.StripLines.Add(sl1)
        Dim sl2 As StripLine = New StripLine()
        sl2.BackColor = Color.Black
        sl2.IntervalOffset = Label18.Text
        sl2.StripWidth = 0.05
        Chart1.ChartAreas(0).AxisY.StripLines.Add(sl2)
    End Sub

    Private Sub AddtoChart()
        Chart1.Series("Height").Points.AddXY(0, h)
        If Chart1.Series("Height").Points.Count = 41 Then
            Chart1.Series("Height").Points.RemoveAt(0)
        End If
    End Sub

    Sub delay(dt As Long)
        Dim tt = GetTickCount
        Do While GetTickCount < dt + tt
            Application.DoEvents()
        Loop
    End Sub

    Private Sub MovePicture()
        Dim fso As Object
        Dim path = Work.workPath + "圖片"
        If Dir(path, vbDirectory) = "" Then
            fso = CreateObject("Scripting.FileSystemObject")
            fso.CreateFolder(path)
            fso = Nothing
        End If
        Shell("cmd" + path)
    End Sub

    Private Sub HeightDataToCsv()
        Try
            Dim cn As New SqlConnection(cnStr)
            cn.Open()
            Dim commandtext As String = ""
            Dim cmd As New SqlCommand(commandtext, cn)
            Dim dr As SqlDataReader
            dr = cmd.ExecuteReader
            Dim dt As DataTable = New DataTable
            dt.Load(dr)
            Dim objExcelApp As Excel.Application
            Dim objSheet As Excel.Worksheet
            Dim objWorkbook As Excel.Workbook
            Dim filePath As String = Work.workPath + ".xlsx"
            objExcelApp = CreateObject("Excel.Application")
            objWorkbook = objExcelApp.Workbooks.Open(filePath)
            objExcelApp.Visible = False
            objSheet = objExcelApp.ActiveWorkbook.Sheets("測高") 
            objSheet.Activate()
            '以下是要填入的資料 
            Dim intColCount As Integer = dt.Columns.Count
            For Each d As DataRow In dt.Rows
                For i As Integer = 0 To intColCount - 4
                    If Not Trim(Convert.ToString(d(i))).Equals("0") Then
                        objSheet.Cells(ht, 1 + i).Value = Trim(Convert.ToString(d(i)))
                    Else
                        objSheet.Cells(ht, 1).Value = "開始時間"
                        objSheet.Cells(ht, 2).Value = Convert.ToString(d(0))
                        objSheet.Cells(ht, 3).Value = "班別"
                        objSheet.Cells(ht, 4).Value = Convert.ToString(d(4))
                        objSheet.Cells(ht, 5).Value = "作業員"
                        objSheet.Cells(ht, 6).Value = "'" + Trim(Convert.ToString(d(5)))
                        Exit For
                    End If
                Next
                ht += 1
            Next
            objExcelApp.DisplayAlerts = False
            objWorkbook.Save()
            objWorkbook.Close(SaveChanges:=True)
            objExcelApp.Quit()
            objExcelApp = Nothing
            objWorkbook = Nothing
            objSheet = Nothing
            GC.Collect()
            Dim commandtext3 As String = "TRUNCATE TABLE HeightTable"
            Dim cmd1 As New SqlCommand(commandtext3, cn)
            cmd1.ExecuteNonQuery()
            cn.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        'End If
    End Sub

    Private Sub OkToCsv()
        Try
            Dim cn As New SqlConnection(cnStr)
            cn.Open()
            Dim commandtext As String = ""
            Dim cmd As New SqlCommand(commandtext, cn)
            Dim dr As SqlDataReader
            dr = cmd.ExecuteReader
            Dim dt As DataTable = New DataTable
            dt.Load(dr)
            Dim objExcelApp As Excel.Application
            Dim objSheet As Excel.Worksheet
            Dim objWorkbook As Excel.Workbook
            Dim filePath As String = Work.workPath + ".xlsx"
            objExcelApp = CreateObject("Excel.Application")
            objWorkbook = objExcelApp.Workbooks.Open(filePath) '路徑 
            objExcelApp.Visible = False
            objSheet = objExcelApp.ActiveWorkbook.Sheets("測高") '表單 
            objSheet.Activate()
            '以下是要填入的資料 
            Dim intColCount As Integer = dt.Columns.Count
            For Each d As DataRow In dt.Rows
                If dt.Columns.Count > 0 AndAlso Not Convert.IsDBNull(d(0)) Then objSheet.Cells(htOK, 7) = (Convert.ToString(d(0)))
                For i As Integer = 0 To intColCount - 1
                    objSheet.Cells(htOK, 7 + i).Value = Convert.ToString(d(i))
                Next
                htOK += 1
            Next
            objExcelApp.DisplayAlerts = False
            objWorkbook.Save()
            objWorkbook.Close(SaveChanges:=True)
            objExcelApp.Quit()
            objExcelApp = Nothing
            objWorkbook = Nothing
            objSheet = Nothing
            GC.Collect()
            Dim commandtext1 As String = ""
            Dim cmd1 As New SqlCommand(commandtext1, cn)
            cmd1.ExecuteNonQuery()
            cn.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub NGToCsv()
        Try
            Dim cn As New SqlConnection(cnStr)
            cn.Open()
            Dim commandtext As String = ""
            Dim cmd As New SqlCommand(commandtext, cn)
            Dim dr As SqlDataReader
            dr = cmd.ExecuteReader
            Dim dt As DataTable = New DataTable
            dt.Load(dr)
            Dim objExcelApp As Excel.Application
            Dim objSheet As Excel.Worksheet
            Dim objWorkbook As Excel.Workbook
            Dim filePath As String = Work.workPath + ".xlsx"
            objExcelApp = CreateObject("Excel.Application")
            objWorkbook = objExcelApp.Workbooks.Open(filePath) 
            objExcelApp.Visible = False
            objSheet = objExcelApp.ActiveWorkbook.Sheets("測高")
            objSheet.Activate()
            '以下是要填入的資料 
            Dim intColCount As Integer = dt.Columns.Count
            For Each d As DataRow In dt.Rows
                For i As Integer = 0 To intColCount - 1
                    objSheet.Cells(htNG, 9 + i).Value = Convert.ToString(d(i))
                Next
                htNG += 1
            Next
            objSheet.Columns.EntireColumn.AutoFit()
            objExcelApp.DisplayAlerts = False
            objWorkbook.Save()
            objWorkbook.Close(SaveChanges:=True)
            objExcelApp.Quit()
            objExcelApp = Nothing
            objWorkbook = Nothing
            objSheet = Nothing
            GC.Collect()
            Dim commandtext1 As String = "TRUNCATE TABLE NGTable"
            Dim cmd1 As New SqlCommand(commandtext1, cn)
            cmd1.ExecuteNonQuery()
            cn.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        ' End If
    End Sub

    Private Sub WriteHeightToCsv()
        HeightDataToCsv()
        OkToCsv()
        NGToCsv()
    End Sub

    Private Sub CCDToCsv()
        Try
            Dim cn As New SqlConnection(cnStr)
            cn.Open()
            Dim commandtext As String = ""
            Dim cmd As New SqlCommand(commandtext, cn)
            Dim dr As SqlDataReader
            dr = cmd.ExecuteReader
            Dim dt As DataTable = New DataTable
            dt.Load(dr)
            Dim objExcelApp As Excel.Application
            Dim objSheet As Excel.Worksheet
            Dim objWorkbook As Excel.Workbook
            Dim filePath As String = Work.workPath + ".xlsx"
            objExcelApp = CreateObject("Excel.Application")
            objWorkbook = objExcelApp.Workbooks.Open(filePath)
            objExcelApp.Visible = False
            objSheet = objExcelApp.ActiveWorkbook.Sheets("工單") 
            objSheet.Activate()
            '以下是要填入的資料 
            objSheet.Cells(3, 2) = Label4.Text
            objSheet.Cells(4, 2) = Label16.Text
            objSheet.Cells(5, 2) = Label15.Text
            objSheet.Cells(6, 2) = Label5.Text
            Dim intColCount As Integer = dt.Columns.Count
            Dim j As Integer = 8
            Dim k As Integer = 0
            For Each d As DataRow In dt.Rows
                objSheet.Cells(j, 0 + 1).Value = "'" + Convert.ToString(d(1))
                For i As Integer = 2 To intColCount - 2
                    objSheet.Cells(j, 0 + i).Value = Convert.ToString(d(i))
                Next
                j = j + 1
                k = k + 1
            Next
            j = j - 1
            For l As Integer = 2 To k
                objSheet.Cells(j, 3).Value = objSheet.Cells(j, 3).Value - objSheet.Cells(j - 1, 3).Value
                objSheet.Cells(j, 4).Value = objSheet.Cells(j, 4).Value - objSheet.Cells(j - 1, 4).Value
                objSheet.Cells(j, 5).Value = objSheet.Cells(j, 5).Value - objSheet.Cells(j - 1, 5).Value
                j = j - 1
            Next
            objSheet.Columns.EntireColumn.AutoFit()
            objExcelApp.DisplayAlerts = False
            objWorkbook.Save()
            objWorkbook.Close(SaveChanges:=True)
            objExcelApp.Quit()
            objExcelApp = Nothing
            objWorkbook = Nothing
            objSheet = Nothing
            GC.Collect()
            Dim commandtext1 As String = "TRUNCATE TABLE CurrentWork"
            Dim cmd1 As New SqlCommand(commandtext1, cn)
            cmd1.ExecuteNonQuery()
            cn.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub LastWork()
        Dim cn As New SqlConnection(cnStr)
        Try
            cn.Open()
            Dim commandtext As String = ""
            Dim cmd As New SqlCommand(commandtext, cn)
            Dim reader As SqlDataReader
            reader = cmd.ExecuteReader
            If reader.HasRows Then
                TPNewWork()
                Dim dt As DataTable = New DataTable
                dt.Load(reader)
                Label1.Text = Trim(dt.Rows(0).Item(0).ToString())
                Label2.Text = Trim(dt.Rows(0).Item(1).ToString())
                Label3.Text = Trim(dt.Rows(0).Item(2).ToString())
                Label5.Text = Trim(dt.Rows(0).Item(3).ToString())
                Label15.Text = Trim(dt.Rows(0).Item(4).ToString())
                Label16.Text = Trim(dt.Rows(0).Item(5).ToString())
                Label4.Text = Trim(Convert.ToInt32(Label5.Text) + Convert.ToInt32(Label15.Text) + Convert.ToInt32(Label16.Text))
                Work.workPath = Trim(dt.Rows(0).Item(7).ToString())
                Button0.Enabled = False
                Button1.Enabled = False
                Button3.Enabled = True
                Button4.Enabled = True
                MsgBox("請結束工單以保存檔案")
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cn.Close()
        End Try
    End Sub

    Private Sub Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If Not Button0.Enabled Then
            e.Cancel = True
            MsgBox("請先結束工單，再關閉程式")
        End If
    End Sub

End Class

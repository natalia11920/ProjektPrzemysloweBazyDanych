Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Data.SqlClient
Imports System.IO



Public Class MainForm

    '*******************************************************************************
    '* Stop polling when the form is not visible in order to reduce communications
    '* Copy this section of code to every new form created
    '*******************************************************************************
    Inherits System.Windows.Forms.Form
    'Create ADO.NET objects.

    Private NotFirstShow As Boolean
    Private Temperature As Double
    Private Humidity As Integer
    Private WindSpeed As Double

    Dim Time As Double = 0
    Dim TimeSQL As Double = 0
    Dim Amplitude As Double = 3
    Dim Counter As Integer = 0
    Dim i As Integer = 1
    Dim WrittenLine As Integer = 0
    Dim K As Integer = 10
    Dim K1 As Integer = 0
    Dim K2 As Integer = 0
    Dim K3 As Integer = 0
    Dim K4 As Integer = 0
    Dim SinusekSQL(1000) As Double
    Dim Sinusek(1000) As Double
    Dim Table As DataTable = New DataTable()
    Dim FileTable As DataTable = New DataTable()
    Dim Flag As Integer = 0
    Dim Flag1 As Integer = 0
    Dim AddSeriesF As Integer = 0
    Dim AddSeriesS As Integer = 0
    Dim h As Integer = 0

    Dim WithEvents UpdateTimer As New Timer()
    Dim WithEvents UpdateTimerSQL As New Timer()
    Dim WithEvents ProcessTimer As New Timer()


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ProcessTimer.Interval = 10000
        ProcessTimer.Start()
    End Sub


    Private Sub ProcessTimer_Tick(sender As Object, e As EventArgs) Handles ProcessTimer.Tick

        Temperature = Rnd() * 40
        Humidity = Rnd() * 100
        WindSpeed = Rnd() * 10

        Temperature = Math.Round(Temperature, 2)
        Humidity = Math.Round(Humidity, 2)
        WindSpeed = Math.Round(WindSpeed, 2)

        CircularProgressBar1.Minimum = 0
        CircularProgressBar1.Maximum = 100

        BarLevel1.Minimum = 0
        BarLevel1.Maximum = 40

        MeterCompact1.Minimum = 0
        MeterCompact1.Maximum = 10

        CircularProgressBar1.Value = Humidity
        BarLevel1.Value = Temperature
        MeterCompact1.Value = WindSpeed


        CircularProgressBar1.Enabled = True
        BarLevel1.Enabled = True
        MeterCompact1.Enabled = True

        If Temperature <= 10 Then
            Me.BarLevel1.BarContentColor = System.Drawing.Color.Blue
        ElseIf Temperature < 20 Then
            Me.BarLevel1.BarContentColor = System.Drawing.Color.LightBlue
        ElseIf Temperature < 30 Then
            Me.BarLevel1.BarContentColor = System.Drawing.Color.LightCoral
        Else
            Me.BarLevel1.BarContentColor = System.Drawing.Color.Crimson
        End If


        Dim timerP As New Stopwatch()
        timerP.Start()

        Dim fileProcess As System.IO.StreamWriter
        fileProcess = My.Computer.FileSystem.OpenTextFileWriter("c:\HMIAdvanced\Measurements.txt", True)
        fileProcess.WriteLine(String.Format($"{Temperature}. {Humidity}. {WindSpeed}"))
        fileProcess.Close()
        timerP.Stop()
        Dim SpeedWriteFile As Long = timerP.ElapsedMilliseconds

        Dim ProcessTable As DataTable = New DataTable()

        'ProcessTable.Clear()
        ProcessTable.Columns.Add("TemperatureM")
        ProcessTable.Columns.Add("HumidityM")
        ProcessTable.Columns.Add("WindSpeedM")

        timerP.Start()
        Dim fileProcessReader As System.IO.StreamReader
        fileProcessReader = My.Computer.FileSystem.OpenTextFileReader("c:\HMIAdvanced\Measurements.txt")


        Do While fileProcessReader.Peek() >= 0
            Dim line As String = fileProcessReader.ReadLine()
            Dim val As String() = line.Split(New String() {". "}, StringSplitOptions.None)
            ProcessTable.Rows.Add(val(0), val(1), val(2))
        Loop

        fileProcessReader.Close()
        timerP.Stop()

        Dim SpeedReadFile As Long = timerP.ElapsedMilliseconds

        h += 1
        timerP.Start()
        Dim Connection As SqlConnection = New SqlConnection("Data Source=LAPTOP-BLHF5NTP\WINCCPLUSMIG2014;Database=Pomiary;Integrated Security=True")
        Dim Command As SqlCommand = New SqlCommand("INSERT INTO Process.dbo.Measurements (ID,Temp,Hum,Wind) VALUES (@ID,@Temperature, @Humidity, @WindS)", Connection)
        Command.Parameters.AddWithValue("@ID", h)
        Command.Parameters.AddWithValue("@Temperature", Temperature)
        Command.Parameters.AddWithValue("@Humidity", Humidity)
        Command.Parameters.AddWithValue("@WindS", WindSpeed)


        Try
            Connection.Open()

            Command.ExecuteNonQuery()
        Catch
            Connection.Close()
        End Try

        timerP.Stop()

        Dim SpeedWriteSQL As Long = timerP.ElapsedMilliseconds

        timerP.Start()

        Dim Connection1 As SqlConnection = New SqlConnection("Data Source=LAPTOP-BLHF5NTP\WINCCPLUSMIG2014;Database=Pomiary;Integrated Security=True")

        Dim SQLReader As SqlDataAdapter = New SqlDataAdapter("SELECT * FROM Process.dbo.Measurements", Connection1)

        Dim TableProcess As DataTable = New DataTable()
        TableProcess.Clear()
        SQLReader.Fill(TableProcess)

        Try
            Connection1.Open()

        Catch
            Connection1.Close()
        End Try

        timerP.Stop()

        Dim SpeedReadSQL As Long = timerP.ElapsedMilliseconds

        OrientedTextLabel10.Text = SpeedWriteFile
        OrientedTextLabel9.Text = SpeedReadFile
        OrientedTextLabel8.Text = SpeedWriteSQL
        OrientedTextLabel7.Text = SpeedReadSQL

        Dim startIndex As Integer = ProcessTable.Rows.Count
        Dim startIndex1 As Integer = TableProcess.Rows.Count

        Dim rowF As DataRow = ProcessTable.Rows(startIndex - 1)
        DataGridView1.Rows.Add(rowF("TemperatureM"), rowF("HumidityM"), rowF("TemperatureM"))

        Dim rowS As DataRow = TableProcess.Rows(startIndex1 - 1)
        DataGridView2.Rows.Add(rowS("Temp"), rowS("Hum"), rowS("Wind"))




    End Sub




    Private Sub Form_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.VisibleChanged
        '* Do not start comms on first show in case it was set to disable in design mode
        If NotFirstShow Then
            AdvancedHMIDrivers.Utilities.StopComsOnHidden(components, Me)
        Else
            NotFirstShow = True
        End If
    End Sub


    Private Sub MainForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Dim index As Integer
        While index < My.Application.OpenForms.Count
            If My.Application.OpenForms(index) IsNot Me Then
                My.Application.OpenForms(index).Close()
            End If
            index += 1
        End While
    End Sub



    Private Sub StartWriting_Click(sender As Object, e As EventArgs) Handles StartWriting.Click

        If Not String.IsNullOrEmpty(TextBox1.Text) Then
            K = TextBox1.Text
            Integer.Parse(K)
        End If

        For L As Integer = 0 To K
            Sinusek(L) = Amplitude * Math.Sin(Time * 2 * Math.PI * 0.5)
            Time += 0.1
        Next

        UpdateTimer.Interval = 5000
        UpdateTimer.Start()



        Dim series As New Series("Sinusoidal Data")
        series.ChartType = SeriesChartType.Line
        If AddSeriesF = 0 Then
            Graph.Series.Add(series)
        End If


        Dim Fseries As New Series("Sinusoidal File Data")
        Fseries.ChartType = SeriesChartType.Line
        If AddSeriesF = 0 Then
            FileGraph.Series.Add(Fseries)
        End If


        Graph.ChartAreas(0).AxisY.Minimum = Amplitude * (-1)
        Graph.ChartAreas(0).AxisY.Maximum = Amplitude * 1



        FileGraph.ChartAreas(0).AxisY.Minimum = Amplitude * (-1)
        FileGraph.ChartAreas(0).AxisY.Maximum = Amplitude * 1


        Dim MaxVisPoint As Integer = 1.5

        If Graph.Series("Sinusoidal Data").Points.Count > MaxVisPoint Then
            Graph.Series("Sinusoidal Data").Points.RemoveAt(0)
        End If

        If FileGraph.Series("Sinusoidal File Data").Points.Count > MaxVisPoint Then
            FileGraph.Series("Sinusoidal File Data").Points.RemoveAt(0)
        End If


    End Sub


    Private Sub UpdateTimer_Tick(sender As Object, e As EventArgs) Handles UpdateTimer.Tick

        Dim TimeInterval As DateTime = DateTime.Now
        Dim format As String = "hh.mm.ss"

        If K3 < K Then
            Graph.Series("Sinusoidal Data").Points.AddXY(TimeInterval.ToString(format), Sinusek(K3))
            Displaying.Text = Sinusek(K3)
            K3 += 1
        End If

        If Flag1 = 0 Then
            FileOperations()
            Flag1 = 1
        End If

        If Flag1 = 1 And K4 < K Then
            Dim SinD As Double
            TimeInterval = DateTime.Now
            format = "hh.mm.ss"
            Dim row2 As DataRow = FileTable.Rows(K4)
            Dim sinusoideText As String = row2("SinusoideValues").ToString()
            Double.TryParse(sinusoideText, SinD)
            FileGraph.Series("Sinusoidal File Data").Points.AddXY(TimeInterval.ToString(format), SinD)
            FileRead.Text = row2("SinusoideValues")
            K4 += 1
        End If

    End Sub


    Private Sub FileOperations()


        Dim timer As New Stopwatch()
        FileTable.Rows.Clear()
        timer.Start()
        Dim file As System.IO.StreamWriter
        file = My.Computer.FileSystem.OpenTextFileWriter("c:\HMIAdvanced\DataSinusoide.txt", True)
        For L As Integer = 0 To K
            file.WriteLine(Sinusek(L))
        Next
        file.Close()
        timer.Stop()

        Dim DurationWriteFile As Long = timer.ElapsedMilliseconds
        Console.WriteLine("Czas zapisywania do pliku: " & DurationWriteFile & " ms")
        OrientedTextLabel21.Text = DurationWriteFile
        If AddSeriesF = 0 Then
            FileTable.Columns.Add("SinusoideValues")
        End If

        timer.Start()

        Dim fileReader As System.IO.StreamReader
        fileReader = My.Computer.FileSystem.OpenTextFileReader("c:\HMIAdvanced\DataSinusoide.txt")


        Do While fileReader.Peek() >= 0
            Dim line As String = fileReader.ReadLine()
            FileTable.Rows.Add(line)
        Loop
        fileReader.Close()
        timer.Stop()

        Dim DurationReadFile As Long = timer.ElapsedMilliseconds
        Console.WriteLine("Czas odczytywania z pliku: " & DurationReadFile & " ms")
        OrientedTextLabel20.Text = DurationReadFile

    End Sub


    'Sinusoida SQL
    Private Sub StartWritingSQL_Click(sender As Object, e As EventArgs) Handles StartWritingSQL.Click

        If Not String.IsNullOrEmpty(TextBox1.Text) Then
            K = TextBox1.Text
            Integer.Parse(K)
        End If


        For L As Integer = 0 To K
            SinusekSQL(L) = Amplitude * Math.Sin(TimeSQL * 2 * Math.PI * 0.5)
            TimeSQL += 0.1
        Next

        UpdateTimerSQL.Interval = 5000
        UpdateTimerSQL.Start()


        Dim Oseries As New Series("Sinusoidal Origin Data")
        Oseries.ChartType = SeriesChartType.Line
        If AddSeriesS = 0 Then
            GraphOrigin.Series.Add(Oseries)
        End If




        Dim Sseries As New Series("Sinusoidal SQL Data")
        Sseries.ChartType = SeriesChartType.Line
        If AddSeriesS = 0 Then
            SQLGraph.Series.Add(Sseries)
        End If



        GraphOrigin.ChartAreas(0).AxisY.Minimum = Amplitude * (-1)
        GraphOrigin.ChartAreas(0).AxisY.Maximum = Amplitude * 1



        SQLGraph.ChartAreas(0).AxisY.Minimum = Amplitude * (-1)
        SQLGraph.ChartAreas(0).AxisY.Maximum = Amplitude * 1


        Dim MaxVisPoint As Integer = 2

        If GraphOrigin.Series("Sinusoidal Origin Data").Points.Count > MaxVisPoint Then
            GraphOrigin.Series("Sinusoidal Origin Data").Points.RemoveAt(0)
        End If

        If SQLGraph.Series("Sinusoidal SQL Data").Points.Count > MaxVisPoint Then
            SQLGraph.Series("Sinusoidal SQL Data").Points.RemoveAt(0)
        End If


    End Sub

    Private Sub UpdateTimerSQL_Tick(sender As Object, e As EventArgs) Handles UpdateTimerSQL.Tick


        Dim TimeInterval1 As DateTime = DateTime.Now
        Dim format1 As String = "hh.mm.ss"

        If K1 < K Then
            GraphOrigin.Series("Sinusoidal Origin Data").Points.AddXY(TimeInterval1.ToString(format1), SinusekSQL(K1))
            Displaying1.Text = SinusekSQL(K1)
            K1 += 1
        End If

        If Flag = 0 Then
            SQLOperations()
            K2 = 0
        End If

        If Flag = 1 And K2 < K Then
            Dim TimeInterval2 As DateTime = DateTime.Now
            Dim format2 As String = "hh.mm.ss"
            Dim row As DataRow = Table.Rows(K2)

            SQLGraph.Series("Sinusoidal SQL Data").Points.AddXY(TimeInterval2.ToString(format2), row("SinusoideValues"))

            SQLServerRead.Text = row("SinusoideValues")
            K2 += 1
        End If

    End Sub



    Private Sub SQLOperations()

        Dim timer_S As New Stopwatch()
        Table.Rows.Clear()
        timer_S.Start()

        For L As Integer = 0 To K
            Dim myConn As SqlConnection = New SqlConnection("Data Source=LAPTOP-BLHF5NTP\WINCCPLUSMIG2014;Database=Pomiary;Integrated Security=True")
            Dim myCmd As SqlCommand = New SqlCommand("INSERT INTO Pomiary.dbo.TableSinusek (Id,SinusoideValues) VALUES (@Id, @SinusoideValues)", myConn)
            myCmd.Parameters.AddWithValue("@Id", L)
            myCmd.Parameters.AddWithValue("@SinusoideValues", SinusekSQL(L))

            Try
                myConn.Open()

                myCmd.ExecuteNonQuery()
            Catch
                myConn.Close()
            End Try
        Next
        timer_S.Stop()

        Dim DurationWriteSQL As Long = timer_S.ElapsedMilliseconds
        Console.WriteLine("Czas wpisywania do bazy SQL: " & DurationWriteSQL & " ms")
        OrientedTextLabel18.Text = DurationWriteSQL

        timer_S.Start()
        For L As Integer = 0 To K
            Dim myConn1 As SqlConnection = New SqlConnection("Data Source=LAPTOP-BLHF5NTP\WINCCPLUSMIG2014;Database=Pomiary;Integrated Security=True")
            Dim Reader As SqlDataAdapter = New SqlDataAdapter("SELECT SinusoideValues FROM Pomiary.dbo.TableSinusek WHERE Id=@Id ", myConn1)
            Reader.SelectCommand.Parameters.AddWithValue("@Id", L)
            Reader.Fill(Table)

            Try
                myConn1.Open()

            Catch
                myConn1.Close()
            End Try
        Next
        timer_S.Stop()

        Dim DurationReadSQL As Long = timer_S.ElapsedMilliseconds
        DurationReadSQL = timer_S.ElapsedMilliseconds
        Console.WriteLine("Czas odczytywania z bazy SQL: " & DurationReadSQL & " ms")
        OrientedTextLabel19.Text = DurationReadSQL

        Flag = 1
    End Sub

    Private Sub Delete_Click(sender As Object, e As EventArgs) Handles Delete.Click
        Dim myConnDelete As SqlConnection = New SqlConnection("Data Source=LAPTOP-BLHF5NTP\WINCCPLUSMIG2014;Database=Pomiary;Integrated Security=True")
        Dim myCmdDelete As SqlCommand = New SqlCommand("DELETE FROM Pomiary.dbo.TableSinusek", myConnDelete)
        Try
            myConnDelete.Open()
            myCmdDelete.ExecuteNonQuery()

        Catch
            myConnDelete.Close()
        End Try

        If File.Exists("c:\HMIAdvanced\DataSinusoide.txt") Then
            My.Computer.FileSystem.DeleteFile("c:\HMIAdvanced\DataSinusoide.txt")
        End If

        Flag = 0
        Flag1 = 0

        UpdateTimerSQL.Stop()
        UpdateTimer.Stop()

        AddSeriesF = 1
        AddSeriesS = 1

        Time = 0
        TimeSQL = 0
        K1 = 0
        K2 = 0
        K3 = 0
        K4 = 0
    End Sub


End Class

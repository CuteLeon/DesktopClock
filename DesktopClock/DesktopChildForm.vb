Imports Microsoft.Win32
Imports System.Net

Public Class DesktopChildForm
    '条件编译指令
#Const Debuging = False
#Const ShowRectangle = False

    '将窗体嵌入桌面
    Private Declare Function SetParent Lib "user32" Alias "SetParent" (ByVal hWndChild As IntPtr, ByVal hWndNewParent As IntPtr) As Integer
    '判断一个窗口句柄是否有效
    Private Declare Function IsWindow Lib "user32" Alias "IsWindow" (ByVal hWnd As IntPtr) As Integer

    Private Const Use24TimeFormat As Boolean = True  '是否使用24小时计时制
    Dim DefaultCityKey As String = "101180101" '默认城市ID（郑州市）
    Dim IntervalDistance As Size = New Size(30, 50) '窗体距离屏幕右上角的距离
    Dim BitmapSize As Size = New Size(800, 420) '位图尺寸
    Dim FormSize As Size = New Size(BitmapSize.Width * (My.Computer.Screen.Bounds.Width / 2732), BitmapSize.Height * (My.Computer.Screen.Bounds.Height / 1536)) '实际显示尺寸（根据屏幕分辨率调整）
    Dim LastMinute As Byte = Now.Minute '上一次记录的分钟数，屏蔽掉无用的工作量
    Dim MonthBitmap As Bitmap = My.Resources.FormResource.ResourceManager.GetObject("Month_" & Now.Month) '月份图
    Dim WeekBitmap As Bitmap = My.Resources.FormResource.ResourceManager.GetObject("Week_" & Now.DayOfWeek) '星期图
    Dim NoonBitmap As Bitmap = My.Resources.FormResource.ResourceManager.GetObject("Noon_" & IIf(Now.Hour > 11, "PM", "AM")) '上下午图
    Dim WeatherBitmap As Bitmap  '天气图
    Dim DesktopIconHandle As IntPtr = GetDesktopIconHandle()

    Private Sub DesktopChildForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        System.Windows.Forms.Control.CheckForIllegalCrossThreadCalls = False
        '调整窗体的尺寸和位置
        Me.Size = FormSize
        Me.Left = My.Computer.Screen.Bounds.Width - Me.Width - IntervalDistance.Width
        Me.Top = IntervalDistance.Height
        '监听外部修改系统时间和日期
        AddHandler SystemEvents.TimeChanged, AddressOf UserChangeTime
        '将窗体设置为桌面图标容器的子窗体，以置后显示
        SetParent(Me.Handle, DesktopIconHandle)
        '启动时初始化界面
        DrawImage(Me, CreateTimeBitmap(GetTimeString()))

#If Not Debuging Then
        '开机自启
        Dim RegStartUp As Microsoft.Win32.RegistryKey = My.Computer.Registry.CurrentUser.CreateSubKey("Software\Microsoft\Windows\CurrentVersion\Run")
        RegStartUp.SetValue("Desktop Clock", Application.ExecutablePath)
        'RegStartUp.DeleteValue("Desktop Clock") '删除开机启动项
#End If

        GetWeather()
    End Sub

    Private Sub UserChangeTime()
        '响应用户在外部修改时间和日期（修改时会连续触发两次此事件）
        LastMinute = Now.Minute '分钟数改变了，记录当前分钟数
        WeekBitmap = My.Resources.FormResource.ResourceManager.GetObject("Week_" & Now.DayOfWeek)
        MonthBitmap = My.Resources.FormResource.ResourceManager.GetObject("Month_" & Now.Month)
        NoonBitmap = My.Resources.FormResource.ResourceManager.GetObject("Noon_" & IIf(Now.Hour > 11, "PM", "AM"))
        DrawImage(Me, CreateTimeBitmap(GetTimeString()))
        GC.Collect()
    End Sub

    Private Sub TimerEngine_Tick(sender As Object, e As EventArgs) Handles TimerEngine.Tick
        If IsWindow(DesktopIconHandle) = False Then
            '桌面句柄被销毁，需要重新查找桌面句柄
            Me.Hide() '暂时隐藏
            DesktopIconHandle = GetDesktopIconHandle()
            If DesktopIconHandle = 0 Then Exit Sub '未找到时退出过程，等待下次查找
            SetParent(Me.Handle, DesktopIconHandle)
            '调整窗体的尺寸和位置
            Me.Left = My.Computer.Screen.Bounds.Width - Me.Width - IntervalDistance.Width
            Me.Top = IntervalDistance.Height
            Me.Show() '重新显示
            'Explorer进程恢复后需要重新刷新界面
            UserChangeTime()
        End If
        If Now.Minute = LastMinute Then Exit Sub
        LastMinute = Now.Minute '分钟数改变了，记录当前分钟数
        If LastMinute = 0 Then '每小时检查一次 上下午、星期和月份，刷新一次当前气温
            WeekBitmap = My.Resources.FormResource.ResourceManager.GetObject("Week_" & Now.DayOfWeek)
            MonthBitmap = My.Resources.FormResource.ResourceManager.GetObject("Month_" & Now.Month)
            NoonBitmap = My.Resources.FormResource.ResourceManager.GetObject("Noon_" & IIf(Now.Hour > 11, "PM", "AM"))
            GetWeather()
        End If
        DrawImage(Me, CreateTimeBitmap(GetTimeString()))
        GC.Collect()
        'Debug.Print("刷新了一次，并进行了资源回收：" & GC.GetTotalMemory(False))
    End Sub

    ''' <summary>根据时间返回相应图像</summary>
    ''' <param name="TimeString">
    ''' 时间参数格式：HHCMMDD
    ''' HH：小时；C：冒号；MM：分钟；DD：日期；
    ''' </param>
    ''' <returns>返回根据时间创建的图像</returns>
    Private Function CreateTimeBitmap(ByVal TimeString As String) As Bitmap
        Dim TimeBitmap As Bitmap = New Bitmap(BitmapSize.Width, BitmapSize.Height)
        Using TimeGraphics As Graphics = Graphics.FromImage(TimeBitmap)
#If ShowRectangle Then
            TimeGraphics.FillRectangle(Brushes.White, 0, 0, BitmapSize.Width, BitmapSize.Height)
#End If
            Dim GraphicsLocationX As Integer = 190 '绘制单个数字的坐标记录
            Dim NumberBitmap As Bitmap = Nothing '单个数字图像
            '得到星期和月份图像
            Dim Index As Integer '字符串内循环因子
            For Index = 0 To 4 '提取时间（前5个字符）
                NumberBitmap = My.Resources.FormResource.ResourceManager.GetObject("Time_" & TimeString.Chars(Index))
#If ShowRectangle Then
                TimeGraphics.FillRectangle(Brushes.Orange, GraphicsLocationX, 0, NumberBitmap.Width, NumberBitmap.Height)
#End If
                TimeGraphics.DrawImage(NumberBitmap, GraphicsLocationX, 0, NumberBitmap.Width, NumberBitmap.Height)
                GraphicsLocationX += NumberBitmap.Width '记录下次绘制的坐标
            Next
            '绘制上下午、月份、星期和分割线
#If ShowRectangle Then
            TimeGraphics.FillRectangle(Brushes.Red, 0, 15, 190, 105)
#End If
            TimeGraphics.DrawImage(NoonBitmap, 0, 15, 190, 105)

#If ShowRectangle Then
            TimeGraphics.FillRectangle(Brushes.Blue, 215, 215, 560, 20)
#End If
            TimeGraphics.DrawImage(My.Resources.FormResource.UI_Tray, 215, 215, 560, 20)

#If ShowRectangle Then
            TimeGraphics.FillRectangle(Brushes.Yellow, 608 - MonthBitmap.Width, 240, MonthBitmap.Width, 86)
#End If
            TimeGraphics.DrawImage(MonthBitmap, 608 - MonthBitmap.Width, 240, MonthBitmap.Width, 86)

#If ShowRectangle Then
            TimeGraphics.FillRectangle(Brushes.Pink, 608 - WeekBitmap.Width, 330, WeekBitmap.Width, 86)
#End If
            TimeGraphics.DrawImage(WeekBitmap, 608 - WeekBitmap.Width, 330, WeekBitmap.Width, 86)
            '绘制日期图像
            GraphicsLocationX = 608
            For Index = 5 To 6
                NumberBitmap = My.Resources.FormResource.ResourceManager.GetObject("Date_" & TimeString.Chars(Index))
#If ShowRectangle Then
                TimeGraphics.FillRectangle(Brushes.Violet, GraphicsLocationX, 248, 96, 162)
#End If
                TimeGraphics.DrawImage(NumberBitmap, GraphicsLocationX, 248, 96, 162)
                GraphicsLocationX += 96
            Next

            If WeatherBitmap IsNot Nothing Then
#If ShowRectangle Then
                TimeGraphics.FillRectangle(Brushes.Navy, 0, 130, 311, 290)
#End If
                TimeGraphics.DrawImage(WeatherBitmap, 0, 130, 311, 290)
            End If

            '释放不需要的内存
            NumberBitmap.Dispose()
            '根据窗体尺寸拉伸时间图像
            TimeBitmap = New Bitmap(TimeBitmap, FormSize)
        End Using
        '返回创建的图像
        Return TimeBitmap
    End Function

    ''' <summary>
    ''' 更新天气信息
    ''' </summary>
    Public Sub GetWeather()
        Try
            If Not My.Computer.Network.Ping("wthrcdn.etouch.cn") Then Exit Sub
        Catch ex As Exception
            Exit Sub
        End Try

        Dim NowTemperature As String, LowTemperature As String, HighTemperature As String, Weather As String
        Dim NowTemperatureClient As WebClient = New WebClient With {.Encoding = System.Text.Encoding.UTF8}
        AddHandler NowTemperatureClient.DownloadDataCompleted, Sub(sender As Object, e As DownloadDataCompletedEventArgs)
                                                                   Dim JsonStrings() As String '= "{""desc"":""OK"",""status"":1000,""data"":{""wendu"":""5"",""ganmao"":""将有一次强降温过程，极易发生感冒，请特别注意增加衣服保暖防寒。"",""forecast"":[{""fengxiang"":""东北风"",""fengli"":""4-5级"",""high"":""高温 6℃"",""type"":""阴"",""low"":""低温 0℃"",""date"":""20日星期一""},{""fengxiang"":""无持续风向"",""fengli"":""微风级"",""high"":""高温 0℃"",""type"":""暴雪"",""low"":""低温 -5℃"",""date"":""21日星期二""},{""fengxiang"":""无持续风向"",""fengli"":""微风级"",""high"":""高温 5℃"",""type"":""阴"",""low"":""低温 -5℃"",""date"":""22日星期三""},{""fengxiang"":""无持续风向"",""fengli"":""微风级"",""high"":""高温 7℃"",""type"":""多云"",""low"":""低温 0℃"",""date"":""23日星期四""},{""fengxiang"":""无持续风向"",""fengli"":""微风级"",""high"":""高温 10℃"",""type"":""晴"",""low"":""低温 0℃"",""date"":""24日星期五""}],""yesterday"":{""fl"":""3-4级"",""fx"":""南风"",""high"":""高温 21℃"",""type"":""晴"",""low"":""低温 6℃"",""date"":""19日星期日""},""aqi"":""142"",""city"":""郑州""}}".Split(New Char() {""""}, 39)
                                                                   '下载的e.Result是经过GZip压缩的，需要解压
                                                                   Using StreamReceived As IO.Stream = New System.IO.MemoryStream(e.Result)
                                                                       Using ZipStream As IO.Compression.GZipStream = New IO.Compression.GZipStream(StreamReceived, IO.Compression.CompressionMode.Decompress)
                                                                           Using DataStreamReader As IO.StreamReader = New IO.StreamReader(ZipStream, System.Text.Encoding.UTF8)
                                                                               Dim JsonString As String = DataStreamReader.ReadToEnd
                                                                               'Debug.Print(JsonString)
                                                                               JsonStrings = Split(JsonString, """", 39)
                                                                           End Using
                                                                       End Using
                                                                   End Using
                                                                   If JsonStrings(3) <> "OK" Then Exit Sub
                                                                   NowTemperature = JsonStrings(11)
                                                                   'JsonStrings(15)：感冒提醒
                                                                   'JsonStrings(21)：风向
                                                                   'JsonStrings(25)：风力
                                                                   HighTemperature = Strings.Mid(JsonStrings(29), 4, JsonStrings(29).Length - 4)
                                                                   Weather = JsonStrings(33)
                                                                   LowTemperature = Strings.Mid(JsonStrings(37), 4, JsonStrings(37).Length - 4)
                                                                   Dim TempWeatherBitmap As Bitmap = New Bitmap(311, 290)
                                                                   Dim DrawLocationX As Integer = 0
                                                                   Using WeatherGraphics As Graphics = Graphics.FromImage(TempWeatherBitmap)
                                                                       Dim WeatherFont As Font, FontSize As Integer = 0, WeatherSize As Size
                                                                       If Weather.Length < 3 Then
                                                                           FontSize = 60
                                                                           WeatherSize = New Size(Weather.Length * 80 + 30, 107)
                                                                       ElseIf Weather.Length = 3 Then
                                                                           FontSize = 45
                                                                           WeatherSize = New Size(200, 85)
                                                                       ElseIf Weather.Length > 3 Then
                                                                           FontSize = 32
                                                                           Weather = Weather.Insert(Weather.Length \ 2, vbCrLf)
                                                                           WeatherSize = New Size(145, 110)
                                                                       End If
                                                                       WeatherFont = New Font(Me.Font.FontFamily, FontSize, FontStyle.Bold)
                                                                       WeatherGraphics.FillRectangle(New SolidBrush(Color.FromArgb(30, 100, 100, 100)), 0, 0, WeatherSize.Width, WeatherSize.Height)
                                                                       WeatherGraphics.DrawString(Weather, WeatherFont, Brushes.White, 0, 0)
                                                                       WeatherFont.Dispose()

                                                                       '绘制当前气温
                                                                       If NowTemperature.StartsWith("-") Then
                                                                           WeatherGraphics.DrawImage(My.Resources.FormResource.Minus, 0, 155, 62, 27)
                                                                           NowTemperature = NowTemperature.Substring(1)
                                                                           DrawLocationX += 62
                                                                       End If
                                                                       For Index As Byte = 0 To NowTemperature.Length - 1
                                                                           WeatherGraphics.DrawImage(My.Resources.FormResource.ResourceManager.GetObject("Temperature_" & NowTemperature.Chars(Index)),
                                                                                                     DrawLocationX, 110, 80, 118)
                                                                           DrawLocationX += 80
                                                                       Next
                                                                       WeatherGraphics.DrawImage(My.Resources.FormResource.Centigrade, DrawLocationX, 110, 89, 89)

                                                                       '绘制最低气温
                                                                       DrawLocationX = 0
                                                                       If LowTemperature.StartsWith("-") Then
                                                                           WeatherGraphics.DrawImage(My.Resources.FormResource.Minus, 0, 255, 30, 13)
                                                                           LowTemperature = LowTemperature.Substring(1)
                                                                           DrawLocationX += 30
                                                                       End If
                                                                       For Index As Byte = 0 To LowTemperature.Length - 1
                                                                           WeatherGraphics.DrawImage(My.Resources.FormResource.ResourceManager.GetObject("Temperature_" & LowTemperature.Chars(Index)),
                                                                                                     DrawLocationX, 231, 36, 53)
                                                                           DrawLocationX += 36
                                                                       Next
                                                                       WeatherGraphics.DrawImage(My.Resources.FormResource.Centigrade, DrawLocationX, 231, 36, 36)

                                                                       DrawLocationX += 36
                                                                       WeatherGraphics.DrawImage(My.Resources.FormResource.Tilde, DrawLocationX, 250, 28, 19)
                                                                       DrawLocationX += 28
                                                                       '绘制最高气温
                                                                       If HighTemperature.StartsWith("-") Then
                                                                           WeatherGraphics.DrawImage(My.Resources.FormResource.Minus, DrawLocationX, 255, 30, 13)
                                                                           HighTemperature = HighTemperature.Substring(1)
                                                                           DrawLocationX += 30
                                                                       End If
                                                                       For Index As Byte = 0 To HighTemperature.Length - 1
                                                                           WeatherGraphics.DrawImage(My.Resources.FormResource.ResourceManager.GetObject("Temperature_" & HighTemperature.Chars(Index)),
                                                                                                     DrawLocationX, 231, 36, 53)
                                                                           DrawLocationX += 36
                                                                       Next
                                                                       WeatherGraphics.DrawImage(My.Resources.FormResource.Centigrade, DrawLocationX, 231, 36, 36)
                                                                   End Using
                                                                   WeatherBitmap = TempWeatherBitmap
                                                                   Invoke(Sub()
                                                                              DrawImage(Me, CreateTimeBitmap(GetTimeString()))
                                                                          End Sub)
                                                               End Sub
        NowTemperatureClient.DownloadDataAsync(New Uri("http://wthrcdn.etouch.cn/weather_mini?citykey=" & DefaultCityKey))
    End Sub



    ''' <summary>
    ''' 产生时间参数
    ''' </summary>
    ''' <param name="Use24TimeFormat">是否使用24小时计时制</param>
    ''' <returns>时间参数（格式：小时C分钟日期）</returns>
    Private Function GetTimeString(Optional ByVal Use24TimeFormat As Boolean = Use24TimeFormat) As String
        Return Now.ToString(IIf(Use24TimeFormat, "HH", "hh") & "Cmmdd")
    End Function

    Protected Overloads Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            If Not DesignMode Then
                Dim cp As CreateParams = MyBase.CreateParams
                '保留窗体原样式的基础上，增加透明 和 鼠标穿透
                cp.ExStyle = cp.ExStyle Or LayeredWindowModule.WS_EX_LAYERED Or LayeredWindowModule.WS_EX_TRANSPARENT
                Return cp
            Else
                Return MyBase.CreateParams
            End If
        End Get
    End Property

End Class

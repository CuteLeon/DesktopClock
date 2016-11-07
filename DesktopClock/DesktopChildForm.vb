Imports Microsoft.Win32
Imports System.Threading

Public Class DesktopChildForm
    '将窗体嵌入桌面
    Private Declare Function SetParent Lib "user32" Alias "SetParent" (ByVal hWndChild As IntPtr, ByVal hWndNewParent As IntPtr) As Integer
    '判断一个窗口句柄是否有效
    Private Declare Function IsWindow Lib "user32" Alias "IsWindow" (ByVal hWnd As IntPtr) As Integer

    Dim IntervalDistance As Size = New Size(30, 50) '窗体距离屏幕右上角的距离
    Dim BitmapSize As Size = New Size(800, 420) '位图尺寸
    Dim FormSize As Size = New Size(400, 210) '窗体显示尺寸（只需要修改这个就可以拉伸数字时钟）
    Dim LastMinute As Byte = Now.Minute '上一次记录的分钟数，屏蔽掉无用的工作量
    Dim WeekBitmap As Bitmap = My.Resources.FormResource.ResourceManager.GetObject("Week_" & Now.DayOfWeek)
    Dim MonthBitmap As Bitmap = My.Resources.FormResource.ResourceManager.GetObject("Month_" & Now.Month)
    Dim NoonBitmap As Bitmap = My.Resources.FormResource.ResourceManager.GetObject("Noon_" & IIf(Now.Hour > 11, "PM", "AM"))
    Dim TrayBitmap As Bitmap = My.Resources.FormResource.UI_Tray
    Dim Use24TimeFormat As Boolean = False '是否使用24小时计时制
    Dim DesktopIconHandle As IntPtr = GetDesktopIconHandle()

    Private Sub DesktopChildForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        '开机自启
        Dim RegStartUp As Microsoft.Win32.RegistryKey = My.Computer.Registry.CurrentUser.CreateSubKey("Software\Microsoft\Windows\CurrentVersion\Run")
        RegStartUp.SetValue("Desktop Clock", Application.ExecutablePath)
        'RegStartUp.DeleteValue("Desktop Clock") '删除开机启动项
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
        End If
        If Now.Minute = LastMinute Then Exit Sub
        LastMinute = Now.Minute '分钟数改变了，记录当前分钟数
        If LastMinute = 0 Then '每小时检查一次 上下午、星期和月份
            WeekBitmap = My.Resources.FormResource.ResourceManager.GetObject("Week_" & Now.DayOfWeek)
            MonthBitmap = My.Resources.FormResource.ResourceManager.GetObject("Month_" & Now.Month)
            NoonBitmap = My.Resources.FormResource.ResourceManager.GetObject("Noon_" & IIf(Now.Hour > 11, "PM", "AM"))
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
        Dim TimeGraphics As Graphics = Graphics.FromImage(TimeBitmap)
        Dim TempChar As Char = vbNullChar '提取自参数的单个字符
        Dim GraphicsLocationX As Integer = 190 '绘制单个数字的坐标记录
        Dim NumberBitmap As Bitmap = Nothing '单个数字图像
        '得到星期和月份图像
        Dim Index As Integer '字符串内循环因子
        For Index = 1 To 5 '提取时间（前5个字符）
            TempChar = Strings.Mid(TimeString, Index, 1)
            NumberBitmap = My.Resources.FormResource.ResourceManager.GetObject("Time_" & TempChar)
            TimeGraphics.DrawImage(NumberBitmap, GraphicsLocationX, 0, NumberBitmap.Width, NumberBitmap.Height)
            GraphicsLocationX += NumberBitmap.Width '记录下次绘制的坐标
        Next
        '绘制上下午、星期和月份图像和分割线
        TimeGraphics.DrawImage(NoonBitmap, 0, 99, 190, 105)
        TimeGraphics.DrawImage(TrayBitmap, 215, 215, 560, 20)
        TimeGraphics.DrawImage(WeekBitmap, 608 - WeekBitmap.Width, 240, WeekBitmap.Width, 90)
        TimeGraphics.DrawImage(MonthBitmap, 608 - MonthBitmap.Width, 330, MonthBitmap.Width, 80)
        '绘制日期图像
        GraphicsLocationX = 608
        For Index = 6 To 7
            TempChar = Strings.Mid(TimeString, Index, 1)
            NumberBitmap = My.Resources.FormResource.ResourceManager.GetObject("Date_" & TempChar)
            TimeGraphics.DrawImage(NumberBitmap, GraphicsLocationX, 248, 96, 162)
            GraphicsLocationX += 96
        Next
        '释放不需要的内存
        NumberBitmap.Dispose()
        '根据窗体尺寸拉伸时间图像
        TimeBitmap = New Bitmap(TimeBitmap, FormSize)
        TimeGraphics.Dispose()
        '返回创建的图像
        Return TimeBitmap
    End Function

    ''' <summary>
    ''' 产生时间参数
    ''' </summary>
    ''' <param name="Use24TimeFormat">是否使用24小时计时制</param>
    ''' <returns>时间参数（格式：小时C分钟日期）</returns>
    Private Function GetTimeString(Optional ByVal Use24TimeFormat As Boolean = False) As String
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

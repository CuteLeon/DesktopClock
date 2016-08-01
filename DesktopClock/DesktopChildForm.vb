Imports Microsoft.Win32
Imports System.Threading

Public Class DesktopChildForm
    Private Declare Function SetParent Lib "user32" Alias "SetParent" (ByVal hWndChild As IntPtr, ByVal hWndNewParent As IntPtr) As Integer
    Dim IntervalDistance As Size = New Size(100, 500)
    Dim BitmapSize As Size = New Size(600, 400)
    Dim FormSize As Size = New Size(300, 200)
    Dim LastMinute As Byte = Now.Minute

    Private Sub DesktopChildForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '调整窗体的尺寸和位置
        Me.Size = FormSize
        Me.Left = (My.Computer.Screen.Bounds.Width - Me.Width) - IntervalDistance.Width
        Me.Top = (My.Computer.Screen.Bounds.Height - Me.Height) - IntervalDistance.Height
        '将窗体设置为桌面图标容器的子窗体，以置后显示
        SetParent(Me.Handle, GetDesktopIconHandle)

        DrawImage(Me, CreateTimeBitmap(Now.Hour.ToString("00") & "C" & LastMinute.ToString("00") & Now.Day.ToString("00")))

        '开机自启
        Dim RegStartUp As Microsoft.Win32.RegistryKey = My.Computer.Registry.CurrentUser.CreateSubKey("Software\Microsoft\Windows\CurrentVersion\Run")
        RegStartUp.SetValue("Desktop Clock", Application.ExecutablePath)
    End Sub

    Private Function CreateTimeBitmap(ByVal TimeString As String)
        '根据时间返回相应图像
        Dim TimeBitmap As Bitmap = New Bitmap(BitmapSize.Width, BitmapSize.Height)
        Dim TimeGraphics As Graphics = Graphics.FromImage(TimeBitmap)
        Dim TempChar As Char = vbNullChar
        Dim GraphicsLocationX As Integer = 0
        Dim NumberBitmap As Bitmap = Nothing
        Dim WeekBitmap As Bitmap = My.Resources.FormResource.ResourceManager.GetObject("Week_" & Now.DayOfWeek)
        Dim MonthBitmap As Bitmap = My.Resources.FormResource.ResourceManager.GetObject("Month_" & Now.Month)
        Dim Index As Integer
        For Index = 1 To 5
            TempChar = Strings.Mid(TimeString, Index, 1)
            NumberBitmap = My.Resources.FormResource.ResourceManager.GetObject("Time_" & TempChar)
            TimeGraphics.DrawImage(NumberBitmap, GraphicsLocationX, 0, NumberBitmap.Width, NumberBitmap.Height)
            GraphicsLocationX += NumberBitmap.Width
        Next
        TimeGraphics.DrawImage(WeekBitmap, 20, 230, WeekBitmap.Width, 90)
        TimeGraphics.DrawImage(MonthBitmap, WeekBitmap.Width - MonthBitmap.Width, 320, MonthBitmap.Width, 80)
        GraphicsLocationX = WeekBitmap.Width + 30
        For Index = 6 To 7
            TempChar = Strings.Mid(TimeString, Index, 1)
            NumberBitmap = My.Resources.FormResource.ResourceManager.GetObject("Date_" & TempChar)
            TimeGraphics.DrawImage(NumberBitmap, GraphicsLocationX, 238, NumberBitmap.Width, NumberBitmap.Height)
            GraphicsLocationX += NumberBitmap.Width
        Next
        NumberBitmap.Dispose()
        TimeBitmap = New Bitmap(TimeBitmap, FormSize)
        TimeGraphics.Dispose()
        Return TimeBitmap
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

    Private Sub TimerEngine_Tick(sender As Object, e As EventArgs) Handles TimerEngine.Tick
        If Now.Minute = LastMinute Then Exit Sub
        LastMinute = Now.Minute
        DrawImage(Me, CreateTimeBitmap(Now.Hour.ToString("00") & "C" & LastMinute.ToString("00") & Now.Day.ToString("00")))
        GC.Collect()
        'Debug.Print("刷新了一次，并进行了资源回收：" & GC.GetTotalMemory(False))
    End Sub
End Class

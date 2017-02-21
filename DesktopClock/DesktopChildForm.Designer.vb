<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class DesktopChildForm
    Inherits System.Windows.Forms.Form

    'Form 重写 Dispose，以清理组件列表。
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows 窗体设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器修改它。  
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.TimerEngine = New System.Windows.Forms.Timer(Me.components)
        Me.SuspendLayout()
        '
        'TimerEngine
        '
        Me.TimerEngine.Enabled = True
        Me.TimerEngine.Interval = 1000
        '
        'DesktopChildForm
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.ClientSize = New System.Drawing.Size(300, 200)
        Me.Font = New System.Drawing.Font("微软雅黑", 26.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "DesktopChildForm"
        Me.Text = "DesktopChildForm"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TimerEngine As Timer
End Class

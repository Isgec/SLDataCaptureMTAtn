<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
    Me.components = New System.ComponentModel.Container()
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmMain))
    Me.gManual = New System.Windows.Forms.GroupBox()
    Me.F_CardNo = New System.Windows.Forms.TextBox()
    Me.cmdConf = New System.Windows.Forms.Button()
    Me.radioProcess = New System.Windows.Forms.RadioButton()
    Me.radioPunch = New System.Windows.Forms.RadioButton()
    Me.F_sdt = New System.Windows.Forms.DateTimePicker()
    Me.F_tdt = New System.Windows.Forms.DateTimePicker()
    Me.cmdStart = New System.Windows.Forms.Button()
    Me.cmdStop = New System.Windows.Forms.Button()
    Me.lbldt = New System.Windows.Forms.Label()
    Me.lbltmc = New System.Windows.Forms.Label()
    Me.lblmc = New System.Windows.Forms.Label()
    Me.lblmsg = New System.Windows.Forms.Label()
    Me.tt = New System.Windows.Forms.ToolTip(Me.components)
    Me.GroupBox1 = New System.Windows.Forms.GroupBox()
    Me.cmdTimerStart = New System.Windows.Forms.Button()
    Me.cmdTimerStop = New System.Windows.Forms.Button()
    Me.gManual.SuspendLayout()
    Me.GroupBox1.SuspendLayout()
    Me.SuspendLayout()
    '
    'gManual
    '
    Me.gManual.Controls.Add(Me.F_CardNo)
    Me.gManual.Controls.Add(Me.cmdConf)
    Me.gManual.Controls.Add(Me.radioProcess)
    Me.gManual.Controls.Add(Me.radioPunch)
    Me.gManual.Controls.Add(Me.F_sdt)
    Me.gManual.Controls.Add(Me.F_tdt)
    Me.gManual.Controls.Add(Me.cmdStart)
    Me.gManual.Controls.Add(Me.cmdStop)
    Me.gManual.ForeColor = System.Drawing.Color.Yellow
    Me.gManual.Location = New System.Drawing.Point(3, 53)
    Me.gManual.Name = "gManual"
    Me.gManual.Size = New System.Drawing.Size(308, 83)
    Me.gManual.TabIndex = 19
    Me.gManual.TabStop = False
    Me.gManual.Text = "Manual"
    '
    'F_CardNo
    '
    Me.F_CardNo.Location = New System.Drawing.Point(199, 19)
    Me.F_CardNo.Name = "F_CardNo"
    Me.F_CardNo.Size = New System.Drawing.Size(62, 20)
    Me.F_CardNo.TabIndex = 10
    Me.tt.SetToolTip(Me.F_CardNo, "Comma separated multiple CardNo")
    '
    'cmdConf
    '
    Me.cmdConf.Cursor = System.Windows.Forms.Cursors.Hand
    Me.cmdConf.FlatStyle = System.Windows.Forms.FlatStyle.Flat
    Me.cmdConf.ForeColor = System.Drawing.Color.Black
    Me.cmdConf.Image = CType(resources.GetObject("cmdConf.Image"), System.Drawing.Image)
    Me.cmdConf.Location = New System.Drawing.Point(267, 12)
    Me.cmdConf.Name = "cmdConf"
    Me.cmdConf.Size = New System.Drawing.Size(30, 30)
    Me.cmdConf.TabIndex = 9
    Me.cmdConf.UseVisualStyleBackColor = True
    '
    'radioProcess
    '
    Me.radioProcess.AutoSize = True
    Me.radioProcess.Location = New System.Drawing.Point(99, 19)
    Me.radioProcess.Name = "radioProcess"
    Me.radioProcess.Size = New System.Drawing.Size(97, 17)
    Me.radioProcess.TabIndex = 8
    Me.radioProcess.Text = "Process Punch"
    Me.radioProcess.UseVisualStyleBackColor = True
    '
    'radioPunch
    '
    Me.radioPunch.AutoSize = True
    Me.radioPunch.Checked = True
    Me.radioPunch.Location = New System.Drawing.Point(13, 19)
    Me.radioPunch.Name = "radioPunch"
    Me.radioPunch.Size = New System.Drawing.Size(81, 17)
    Me.radioPunch.TabIndex = 7
    Me.radioPunch.TabStop = True
    Me.radioPunch.Text = "Raw Punch"
    Me.radioPunch.UseVisualStyleBackColor = True
    '
    'F_sdt
    '
    Me.F_sdt.CustomFormat = "dd/MM/yy"
    Me.F_sdt.Format = System.Windows.Forms.DateTimePickerFormat.Custom
    Me.F_sdt.Location = New System.Drawing.Point(9, 46)
    Me.F_sdt.Name = "F_sdt"
    Me.F_sdt.Size = New System.Drawing.Size(75, 20)
    Me.F_sdt.TabIndex = 1
    '
    'F_tdt
    '
    Me.F_tdt.CustomFormat = "dd/MM/yy"
    Me.F_tdt.Format = System.Windows.Forms.DateTimePickerFormat.Custom
    Me.F_tdt.Location = New System.Drawing.Point(90, 46)
    Me.F_tdt.Name = "F_tdt"
    Me.F_tdt.Size = New System.Drawing.Size(72, 20)
    Me.F_tdt.TabIndex = 3
    '
    'cmdStart
    '
    Me.cmdStart.FlatAppearance.BorderColor = System.Drawing.Color.Red
    Me.cmdStart.FlatAppearance.CheckedBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
    Me.cmdStart.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
    Me.cmdStart.FlatAppearance.MouseOverBackColor = System.Drawing.Color.CornflowerBlue
    Me.cmdStart.FlatStyle = System.Windows.Forms.FlatStyle.Flat
    Me.cmdStart.Location = New System.Drawing.Point(168, 47)
    Me.cmdStart.Name = "cmdStart"
    Me.cmdStart.Size = New System.Drawing.Size(60, 23)
    Me.cmdStart.TabIndex = 5
    Me.cmdStart.Text = "Start"
    Me.cmdStart.UseVisualStyleBackColor = True
    '
    'cmdStop
    '
    Me.cmdStop.Enabled = False
    Me.cmdStop.FlatAppearance.BorderColor = System.Drawing.Color.Red
    Me.cmdStop.FlatAppearance.CheckedBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
    Me.cmdStop.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
    Me.cmdStop.FlatAppearance.MouseOverBackColor = System.Drawing.Color.CornflowerBlue
    Me.cmdStop.FlatStyle = System.Windows.Forms.FlatStyle.Flat
    Me.cmdStop.Location = New System.Drawing.Point(234, 47)
    Me.cmdStop.Name = "cmdStop"
    Me.cmdStop.Size = New System.Drawing.Size(60, 23)
    Me.cmdStop.TabIndex = 6
    Me.cmdStop.Text = "Stop"
    Me.cmdStop.UseVisualStyleBackColor = True
    '
    'lbldt
    '
    Me.lbldt.AutoSize = True
    Me.lbldt.ForeColor = System.Drawing.Color.LawnGreen
    Me.lbldt.Location = New System.Drawing.Point(9, 9)
    Me.lbldt.Name = "lbldt"
    Me.lbldt.Size = New System.Drawing.Size(85, 13)
    Me.lbldt.TabIndex = 20
    Me.lbldt.Text = "Processing Date"
    '
    'lbltmc
    '
    Me.lbltmc.AutoSize = True
    Me.lbltmc.ForeColor = System.Drawing.Color.Aqua
    Me.lbltmc.Location = New System.Drawing.Point(103, 9)
    Me.lbltmc.Name = "lbltmc"
    Me.lbltmc.Size = New System.Drawing.Size(80, 13)
    Me.lbltmc.TabIndex = 21
    Me.lbltmc.Text = "Total Machines"
    '
    'lblmc
    '
    Me.lblmc.AutoSize = True
    Me.lblmc.ForeColor = System.Drawing.Color.LightPink
    Me.lblmc.Location = New System.Drawing.Point(199, 9)
    Me.lblmc.Name = "lblmc"
    Me.lblmc.Size = New System.Drawing.Size(103, 13)
    Me.lblmc.TabIndex = 22
    Me.lblmc.Text = "Connected Machine"
    '
    'lblmsg
    '
    Me.lblmsg.ForeColor = System.Drawing.Color.Fuchsia
    Me.lblmsg.Location = New System.Drawing.Point(8, 26)
    Me.lblmsg.Name = "lblmsg"
    Me.lblmsg.Size = New System.Drawing.Size(302, 27)
    Me.lblmsg.TabIndex = 23
    Me.lblmsg.Text = "Processing Message"
    '
    'GroupBox1
    '
    Me.GroupBox1.Controls.Add(Me.cmdTimerStart)
    Me.GroupBox1.Controls.Add(Me.cmdTimerStop)
    Me.GroupBox1.ForeColor = System.Drawing.Color.Yellow
    Me.GroupBox1.Location = New System.Drawing.Point(317, 53)
    Me.GroupBox1.Name = "GroupBox1"
    Me.GroupBox1.Size = New System.Drawing.Size(84, 83)
    Me.GroupBox1.TabIndex = 24
    Me.GroupBox1.TabStop = False
    Me.GroupBox1.Text = "Scheduler"
    '
    'cmdTimerStart
    '
    Me.cmdTimerStart.FlatAppearance.BorderColor = System.Drawing.Color.Red
    Me.cmdTimerStart.FlatAppearance.CheckedBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
    Me.cmdTimerStart.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
    Me.cmdTimerStart.FlatAppearance.MouseOverBackColor = System.Drawing.Color.CornflowerBlue
    Me.cmdTimerStart.FlatStyle = System.Windows.Forms.FlatStyle.Flat
    Me.cmdTimerStart.Location = New System.Drawing.Point(10, 20)
    Me.cmdTimerStart.Name = "cmdTimerStart"
    Me.cmdTimerStart.Size = New System.Drawing.Size(60, 23)
    Me.cmdTimerStart.TabIndex = 5
    Me.cmdTimerStart.Text = "Start"
    Me.cmdTimerStart.UseVisualStyleBackColor = True
    '
    'cmdTimerStop
    '
    Me.cmdTimerStop.Enabled = False
    Me.cmdTimerStop.FlatAppearance.BorderColor = System.Drawing.Color.Red
    Me.cmdTimerStop.FlatAppearance.CheckedBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
    Me.cmdTimerStop.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
    Me.cmdTimerStop.FlatAppearance.MouseOverBackColor = System.Drawing.Color.CornflowerBlue
    Me.cmdTimerStop.FlatStyle = System.Windows.Forms.FlatStyle.Flat
    Me.cmdTimerStop.Location = New System.Drawing.Point(10, 49)
    Me.cmdTimerStop.Name = "cmdTimerStop"
    Me.cmdTimerStop.Size = New System.Drawing.Size(60, 23)
    Me.cmdTimerStop.TabIndex = 6
    Me.cmdTimerStop.Text = "Stop"
    Me.cmdTimerStop.UseVisualStyleBackColor = True
    '
    'FrmMain
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.BackColor = System.Drawing.Color.Black
    Me.ClientSize = New System.Drawing.Size(410, 144)
    Me.Controls.Add(Me.GroupBox1)
    Me.Controls.Add(Me.lblmsg)
    Me.Controls.Add(Me.lblmc)
    Me.Controls.Add(Me.lbltmc)
    Me.Controls.Add(Me.lbldt)
    Me.Controls.Add(Me.gManual)
    Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.MaximizeBox = False
    Me.MinimizeBox = False
    Me.Name = "FrmMain"
    Me.Text = "Data Capture"
    Me.gManual.ResumeLayout(False)
    Me.gManual.PerformLayout()
    Me.GroupBox1.ResumeLayout(False)
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents gManual As System.Windows.Forms.GroupBox
  Friend WithEvents F_sdt As System.Windows.Forms.DateTimePicker
  Friend WithEvents F_tdt As System.Windows.Forms.DateTimePicker
  Friend WithEvents cmdStart As System.Windows.Forms.Button
  Friend WithEvents cmdStop As System.Windows.Forms.Button
  Friend WithEvents lbldt As System.Windows.Forms.Label
  Friend WithEvents lbltmc As System.Windows.Forms.Label
  Friend WithEvents lblmc As System.Windows.Forms.Label
  Friend WithEvents lblmsg As System.Windows.Forms.Label
  Friend WithEvents radioProcess As System.Windows.Forms.RadioButton
  Friend WithEvents radioPunch As System.Windows.Forms.RadioButton
  Friend WithEvents cmdConf As System.Windows.Forms.Button
  Friend WithEvents F_CardNo As System.Windows.Forms.TextBox
  Friend WithEvents tt As System.Windows.Forms.ToolTip
  Friend WithEvents GroupBox1 As GroupBox
  Friend WithEvents cmdTimerStart As Button
  Friend WithEvents cmdTimerStop As Button
End Class

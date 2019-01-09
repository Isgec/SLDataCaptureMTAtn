Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Imports Isgec_StarLink
Imports AttendanceProcess
Public Class FrmMain
  Private Delegate Sub Attendance()
  Private Delegate Sub Punch()
  Private frmClosing As Boolean = False
  Private Processing As Boolean = False
  Private mayStop As Boolean = False
  Dim DataFilePath As String = Application.StartupPath
  Dim mcList As New ArrayList
  Dim LogWriter As IO.StreamWriter
  Dim conStr As String = ""
  Private smtpServer As String = ""
  Private fromEmail As String = ""
  Private fromID As String = ""
  Private fromPW As String = ""
  Private cc As String = ""
  Private Sub msg(ByVal str As String)
    lblmsg.Text = str
    Try
      LogWriter.WriteLine(str)
    Catch ex As Exception
    End Try
  End Sub
  Private Sub FrmMain_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
    frmClosing = True
    Select Case e.CloseReason
      Case CloseReason.FormOwnerClosing, CloseReason.MdiFormClosing, CloseReason.TaskManagerClosing, CloseReason.UserClosing, CloseReason.WindowsShutDown
        If Processing Then
          mayStop = True
          e.Cancel = True
          msg("Closing Application Wait...")
        End If
    End Select
  End Sub
  Private Sub FrmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    CheckForIllegalCrossThreadCalls = False
    Try
      If modMain.AutoStart Then
        Me.Height = 82
        Me.gManual.Visible = False
      Else
        Me.Height = 167
        Me.gManual.Visible = True
      End If
      Me.F_sdt.Value = Now
      Me.F_tdt.Value = Now
      Dim ts As IO.StreamReader = New IO.StreamReader(Application.StartupPath & "\IsgecMC.config")
      Dim II As Integer = 0
      Dim tmp As String = ts.ReadLine
      Do While tmp IsNot Nothing
        If Not tmp.StartsWith("#") Then
          Dim aTmp As String() = tmp.Split("|".ToCharArray)
          Try
            Select Case aTmp(0).ToLower
              Case "data file path"
                DataFilePath = aTmp(1)
              Case "connection string"
                conStr = aTmp(1)
              Case "SMTP Server".ToLower
                smtpServer = aTmp(1)
              Case "From EMail".ToLower
                fromEmail = aTmp(1)
              Case "From ID".ToLower
                fromID = aTmp(1)
              Case "From PW".ToLower
                fromPW = aTmp(1)
              Case "CC".ToLower
                cc = aTmp(1)
              Case "punch machine"
                mcList.Add(aTmp(1))
            End Select
          Catch ex As Exception
          End Try
        End If
        tmp = ts.ReadLine
      Loop
      ts.Close()
      Me.lbltmc.Text = mcList.Count
      If modMain.AutoStart Or modMain.AutoProcess Or modMain.AutoProcessLastDate Then
        If modMain.AutoStart Then Me.radioPunch.Checked = True
        If modMain.AutoProcess Or modMain.AutoProcessLastDate Then Me.radioProcess.Checked = True
        If modMain.AutoProcessLastDate Then
          Me.F_sdt.Value = Now.AddDays(-1)
          Me.F_tdt.Value = Now.AddDays(-1)
        End If
        cmdStart_Click(Nothing, Nothing)
      End If
    Catch ex As Exception
      msg("File: IsgecMC.config NOT Found. Can NOT Continue.")
      Try
        Dim ts As IO.StreamWriter = New IO.StreamWriter(Application.StartupPath & "\IsgecMC.config")
        ts.WriteLine("#====================================================")
        ts.WriteLine("#Sample IsgecMC.config File")
        ts.WriteLine("#-----------------------------------------------------")
        ts.WriteLine("#Connection String| To Connect SQL Data Base")
        ts.WriteLine("#Table Name|")
        ts.WriteLine("#Time Field|")
        ts.WriteLine("#Card No Field|")
        ts.WriteLine("#Date Field|")
        ts.WriteLine("#Data File Path|To Create TEXT file of downloaded data")
        ts.WriteLine("#SMTP Server IP")
        ts.WriteLine("#From E-Mail ID")
        ts.WriteLine("#EMail Client ID")
        ts.WriteLine("#EMail Client Password")
        ts.WriteLine("#CC E-Mail ID-to store sent mail")
        ts.WriteLine("#Punch Machine|Punch Machine Name, IP Address")
        ts.WriteLine("#======================================================")
        ts.WriteLine("Connection String|Data Source=192.9.200.169;Initial Catalog=ISGEC;Integrated Security=False;User Instance=False;Persist Security Info=True;User ID=sa;Password=Webpay@2013;Connection Timeout=900")
        ts.WriteLine("Data File Path|C:\Temp")
        ts.WriteLine("SMTP Server|192.9.200.214")
        ts.WriteLine("From EMail|leave@isgec.co.in")
        ts.WriteLine("From ID|leave")
        ts.WriteLine("From PW|ijt123")
        ts.WriteLine("CC|tsld@isgec.co.in")
        ts.WriteLine("Punch Machine|01-A4-03,192.168.28.11")
        ts.WriteLine("Punch Machine|01-A4-04,192.168.28.15")
        ts.WriteLine("Punch Machine|01-A4-05,192.168.28.10")
        ts.WriteLine("Punch Machine|01-A4-06,192.168.28.4")
        ts.WriteLine("Punch Machine|02-A56-03,192.168.17.11")
        ts.WriteLine("Punch Machine|02-A56-04,192.168.17.12")
        ts.WriteLine("Punch Machine|04-CHENNAI-02,192.168.33.250")
        ts.WriteLine("Punch Machine|04-CHENNAI-03,192.168.33.251")
        ts.WriteLine("Punch Machine|05-PUNE-01,192.168.10.251")
        ts.WriteLine("Punch Machine|10-A7-03,192.168.25.18")
        ts.WriteLine("Punch Machine|10-A7-04,192.168.25.15")
        ts.WriteLine("Punch Machine|10-A7-05,192.168.24.21")
        ts.WriteLine("Punch Machine|14-A8,192.168.8.2")
        ts.WriteLine("Punch Machine|11-A5-01,192.168.5.20")
        ts.WriteLine("Punch Machine|11-A5-02,192.168.5.21")
        ts.Close()
      Catch ex1 As Exception
      End Try
    End Try
  End Sub
  Private Sub cmdStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStart.Click
    Processing = True
    mayStop = False
    cmdStart.Enabled = False
    cmdStop.Enabled = True
    If Me.radioPunch.Checked Then
      If modMain.AutoStart Then
        Dim DataOK As IO.StreamWriter = New IO.StreamWriter(DataFilePath & "\DataOK.lg")
        DataOK.WriteLine("Started Downloading at " & Now)
        DataOK.Close()
      End If
      RawPunch.UpdatePunch.ConnectionString = conStr
      Dim dp As Punch = AddressOf DoPunch
      dp.BeginInvoke(Nothing, Nothing)
    Else
      If modMain.AutoProcess Or modMain.AutoProcessLastDate Then
        Try
          Dim DataOK As IO.StreamReader = New IO.StreamReader(DataFilePath & "\DataOK.lg")
          If DataOK.ReadLine().StartsWith("Completed") Then
            DataOK.Close()
            Dim pFail As IO.StreamWriter = New IO.StreamWriter(DataFilePath & "\DataOK.lg")
            pFail.WriteLine("Starting Processing But Last Download NOT completed." & Now)
            pFail.Close()
          End If
        Catch ex As Exception
        End Try
      End If
      AttendanceProcess.NewAttendanceRules.ConnectionString = conStr
      AttendanceProcess.NewAttendanceRules.smtpServer = smtpServer
      AttendanceProcess.NewAttendanceRules.fromEmail = fromEmail
      AttendanceProcess.NewAttendanceRules.fromID = fromID
      AttendanceProcess.NewAttendanceRules.fromPW = fromPW
      AttendanceProcess.NewAttendanceRules.cc = cc
      Dim pdp As Attendance = AddressOf DoProcess
      pdp.BeginInvoke(Nothing, Nothing)
    End If
  End Sub
  Private Sub DoProcess()
    Try
      If DataFilePath <> String.Empty Then
        LogWriter = New IO.StreamWriter(DataFilePath & "\" & Now.Year & Now.Month.ToString.PadLeft(2, "0") & Now.Day.ToString.PadLeft(2, "0") & Now.Hour.ToString.PadLeft(2, "0") & Now.Minute.ToString.PadLeft(2, "0") & ".log")
      End If
    Catch ex As Exception
    End Try
    AttendanceProcess.NewAttendanceRules.MayStopProcess = False
    Try
      If AutoProcessLastDate Then
        If F_CardNo.Text = String.Empty Then
          ProcessData(Now.AddDays(-1).ToString("dd/MM/yyyy"), Now.AddDays(-1).ToString("dd/MM/yyyy"))
        Else
          ProcessEmpData(Now.AddDays(-1).ToString("dd/MM/yyyy"), Now.AddDays(-1).ToString("dd/MM/yyyy"), F_CardNo.Text)
        End If
      Else
        If F_CardNo.Text = String.Empty Then
          ProcessData(F_sdt.Text, F_sdt.Text)
        Else
          ProcessEmpData(F_sdt.Text, F_sdt.Text, F_CardNo.Text)
        End If
      End If
    Catch ex As Exception
      msg(ex.Message)
    End Try
    cmdStop.Enabled = False
    cmdStart.Enabled = True
    Processing = False
    Try
      LogWriter.Close()
    Catch ex As Exception
    End Try
    If frmClosing Then
      Application.Exit()
    ElseIf modMain.AutoProcess Or modMain.AutoProcessLastDate Then
      Application.Exit()
    End If
  End Sub
  Private Function ProcessData(ByVal FromDate As String, ByVal ToDate As String) As String
    Dim ci As System.Globalization.CultureInfo = New System.Globalization.CultureInfo("en-GB", True)
    Dim fdt As DateTime = Convert.ToDateTime(FromDate, ci)
    Dim tdt As DateTime = Convert.ToDateTime(ToDate, ci)
    Dim xdt As DateTime = fdt
    'Without Employee Code
    Dim StartRow As Integer = 0
    Dim MaxRow As Integer = 10
    Dim oEmps As List(Of SIS.ATN.atnEmployees) = SIS.ATN.atnEmployees.SelectList(StartRow, MaxRow, "CardNo", False, "")
    Me.lbltmc.Text = SIS.ATN.atnEmployees.SelectCount(False, "")
    Do While oEmps.Count > 0
      Try
        Me.lblmc.Text = StartRow
        For Each oEmp As SIS.ATN.atnEmployees In oEmps
          Try
            Me.lblmc.Text = Convert.ToInt32(Me.lblmc.Text) + 1
            'C_OfficeID is blank then skip
            If oEmp.C_OfficeID = "" Then
              msg(xdt.ToString("dd/MM/yyyy") & ", " & oEmp.CardNo & ", " & "Not Verified, Current Location is Blank")
              Continue For
            End If
            Try
              xdt = fdt
              While xdt <= tdt
                Try
                  Me.lbldt.Text = xdt
                  AttendanceProcess.NewAttendanceRules.FinYear = Year(xdt)
                  msg(oEmp.CardNo & " : " & oEmp.EmployeeName & " : " & xdt.ToString("dd/MM/yyyy") & "=>" & AttendanceProcess.NewAttendanceRules.ActualProcess(xdt, oEmp))
                  xdt = xdt.AddDays(1)
                Catch ex As Exception
                End Try
              End While
              'Process Advance Application
              Try
                AttendanceProcess.NewAttendanceRules.PostUnpostedAdavanceApplication(tdt, oEmp.CardNo)
              Catch ex As Exception
                msg("Post Unposted Advance Application Error: " & ex.Message)
              End Try
            Catch ex As Exception
              msg(xdt.ToString("dd/MM/yyyy") & ", " & oEmp.CardNo & ", " & ex.Message)
            End Try
            If mayStop Then
              Exit For
            End If
          Catch ex As Exception
          End Try
        Next
        If mayStop Then
          Exit Do
        End If
        StartRow += MaxRow
        oEmps = SIS.ATN.atnEmployees.SelectList(StartRow, MaxRow, "CardNo", False, "")
      Catch ex As Exception
      End Try
    Loop
    msg("***!!! Process Over !!!***")
    Return ""
  End Function
  Private Function ProcessEmpData(ByVal FromDate As String, ByVal ToDate As String, Optional ByVal CardNo As String = "") As String
    Dim ci As System.Globalization.CultureInfo = New System.Globalization.CultureInfo("en-GB", True)
    Dim fdt As DateTime = Convert.ToDateTime(FromDate, ci)
    Dim tdt As DateTime = Convert.ToDateTime(ToDate, ci)
    Dim xdt As DateTime = fdt
    'Without Employee Code
    Dim StartRow As Integer = 0
    Dim MaxRow As Integer = 10
    Dim aCardNo() As String = CardNo.Split(",".ToCharArray)
    Me.lbltmc.Text = aCardNo.Length
    For Each card As String In aCardNo
      Try
        Dim oEmp As SIS.ATN.atnEmployees = SIS.ATN.atnEmployees.GetByID(card)
        StartRow += 1
        Me.lblmc.Text = StartRow

        Me.lblmc.Text = Convert.ToInt32(Me.lblmc.Text) + 1
        'C_OfficeID is blank then skip
        If oEmp.C_OfficeID = "" Then
          msg(xdt.ToString("dd/MM/yyyy") & ", " & oEmp.CardNo & ", " & "Not Verified, Current Location is Blank")
          Continue For
        End If
        Try
          xdt = fdt
          While xdt <= tdt
            Try
              Me.lbldt.Text = xdt
              AttendanceProcess.NewAttendanceRules.FinYear = Year(xdt)
              msg(oEmp.CardNo & " : " & oEmp.EmployeeName & " : " & xdt.ToString("dd/MM/yyyy") & "=>" & AttendanceProcess.NewAttendanceRules.ActualProcess(xdt, oEmp))
              xdt = xdt.AddDays(1)
            Catch ex As Exception
            End Try
          End While
          'Process Advance Application
          Try
            AttendanceProcess.NewAttendanceRules.PostUnpostedAdavanceApplication(tdt, oEmp.CardNo)
          Catch ex As Exception
            msg("Post Unposted Advance Application Error: " & ex.Message)
          End Try
        Catch ex As Exception
          msg(xdt.ToString("dd/MM/yyyy") & ", " & oEmp.CardNo & ", " & ex.Message)
        End Try
        If mayStop Then
          Exit For
        End If
      Catch ex As Exception
      End Try
    Next
    msg("***!!! Process Over !!!***")
    Return ""
  End Function

  Private Sub CreateTextFile(ByVal FileName As String, ByVal s As String)
    Dim ts As IO.StreamWriter = Nothing
    Try
      ts = New IO.StreamWriter(FileName)
    Catch ex As Exception
      Exit Sub
    End Try
    Dim aStr() As String = s.Split(Chr(10) & Chr(13))
    Dim dat As String = ""
    Dim cardno As String = ""
    For I As Integer = 1 To aStr.Length - 2
      Dim tmp As String = aStr(I).Trim
      cardno = Trim(Mid(tmp, 5, 8))
      dat = Mid(tmp, 1, 2) & ":" & Mid(tmp, 3, 2)
      ts.WriteLine(cardno.PadLeft(9, " ") & " " & dat & "       I")
    Next
    ts.Close()
  End Sub
  Private Sub DoPunch()
    Dim MachineName As Integer = 0
    Dim MachineIP As String = 1
    Dim cnt As Integer = 0
    Dim Err As Boolean = False
    Try
      If DataFilePath <> String.Empty Then
        LogWriter = New IO.StreamWriter(DataFilePath & "\" & Now.Year & Now.Month.ToString.PadLeft(2, "0") & Now.Day.ToString.PadLeft(2, "0") & Now.Hour.ToString.PadLeft(2, "0") & Now.Minute.ToString.PadLeft(2, "0") & ".log")
      End If
    Catch ex As Exception
      Err = True
    End Try
    Dim oCn As New Isgec_StarLink.clsDevice
    '1. Process Each Machine
    For Each itm As String In mcList
      cnt += 1
      Me.lblmc.Text = cnt
      Dim PunchMC() As String = itm.Split(",".ToCharArray)
      msg("Connecting Punch Machine: " & PunchMC(MachineName))
      Dim ret As Integer = oCn.Connect(PunchMC(MachineIP), 1085)
      If ret = 0 Then
        msg("Connected to Punch Machine: " & PunchMC(MachineName))
        Dim pData As String = ""
        Dim sdt As String = ""
        Dim tdt As String = ""
        sdt = F_sdt.Text.Replace("/", "")
        tdt = F_tdt.Text.Replace("/", "")
        Dim ftmp As Date = CDate(Mid(sdt, 1, 2) & "/" & Mid(sdt, 3, 2) & "/20" & Mid(sdt, 5, 2))
        Dim ttmp As Date = CDate(Mid(tdt, 1, 2) & "/" & Mid(tdt, 3, 2) & "/20" & Mid(tdt, 5, 2))
        '2. Process for Each Date in Range
        Do While ftmp <= ttmp
          Me.lbldt.Text = ftmp
          msg("Getting Data for: " & ftmp.ToString("dd/MM/yyyy"))
          sdt = ftmp.ToString("ddMMyy")
          pData = oCn.GetPunchData(sdt)
          If DataFilePath <> String.Empty Then
            Try
              CreateTextFile(DataFilePath & "\" & PunchMC(MachineName).Trim & sdt & ".txt", pData)
            Catch ex As Exception
              Err = True
            End Try
          End If
          ListData(UpdateSQL(pData, ftmp, PunchMC(MachineName)))
          If mayStop Then
            msg("***Cancelled by user***")
            Err = True
            Exit Do
          End If
          ftmp = DateAdd("d", 1, ftmp)
        Loop
        If mayStop Then
          oCn.Disconnect()
          Exit For
        End If
      Else
        msg("Error Could NOT connect to Punch Machine: " & PunchMC(MachineName))
        Err = True
      End If
      If mayStop Then
        Exit For
      End If
      If ret = 0 Then
        msg("Disconnecting " & PunchMC(MachineName))
        Try
          ret = oCn.Disconnect()
          msg("Disconnected " & PunchMC(MachineName))
        Catch ex As Exception
        End Try
      End If
    Next
    If Not mayStop Then
      msg("!!!!!*****Process Over*****!!!!!")
    End If
    If Not Err Then
      Dim DataOK As IO.StreamWriter = New IO.StreamWriter(DataFilePath & "\DataOK.lg")
      DataOK.WriteLine("Completed Downloading at " & Now)
      DataOK.Close()
    End If
    oCn = Nothing
    cmdStop.Enabled = False
    cmdStart.Enabled = True
    Processing = False
    Try
      LogWriter.Close()
    Catch ex As Exception
    End Try
    If frmClosing Then
      Application.Exit()
    ElseIf modMain.AutoStart Then
      Application.Exit()
    End If
  End Sub
  Private Sub ListData(ByVal s As String())
    For Each t As String In s
      msg(t)
    Next
  End Sub
  Private Sub cmdStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStop.Click
    msg("Cancelling. . . ")
    mayStop = True
    AttendanceProcess.NewAttendanceRules.MayStopProcess = True
    cmdStop.Enabled = False
  End Sub
  Private Sub F_sdt_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    F_tdt.Text = F_sdt.Text
  End Sub
  Private Sub F_tdt_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    If Not Convert.ToDateTime(F_tdt.Text) >= Convert.ToDateTime(F_sdt.Text) Then
      F_sdt.Text = F_tdt.Text
    End If
  End Sub

  Public Function UpdateSQL(ByVal s As String, ByVal DataDate As DateTime, ByVal MachineName As String) As String()
    If s Is Nothing Then
      Return {"No Data Found"}
    ElseIf s.Trim = String.Empty Then
      Return {"No Data Found"}
    Else
      Try
        Dim t As Int64 = Convert.ToInt64(s.Trim)
        If t <= 0 Then
          Return {"No Data Found"}
        End If
      Catch ex As Exception
      End Try
    End If
    Dim aStr() As String
    Dim PunchTime As String
    Dim CardNo As String
    aStr = s.Split(Chr(10) & Chr(13))
    For i As Integer = 1 To aStr.Length - 2
      aStr(i) = aStr(i).Trim
      Dim tmp As String = aStr(i)
      CardNo = Trim(Mid(tmp, 5, 8))
      PunchTime = Mid(tmp, 1, 2) & ":" & Mid(tmp, 3, 2)
      Try
        RawPunch.UpdatePunch.FinYear = DataDate.Year
        RawPunch.UpdatePunch.FileUnderProcess = MachineName
        RawPunch.UpdatePunch.UpdateRawData(CardNo, PunchTime, DataDate)
        aStr(i) &= " Updated"
        msg(aStr(i))
      Catch ex As Exception
        aStr(i) &= " Error " & ex.Message & vbCrLf
        msg(aStr(i))
      End Try
      If mayStop Then
        For j = i + 1 To aStr.Length - 2
          aStr(j) = aStr(j).Trim & " Cancelled by User"
          msg(aStr(j))
        Next
        Exit For
      End If
    Next
    Return {""}
  End Function

  Private Sub cmdConf_Click(sender As System.Object, e As System.EventArgs) Handles cmdConf.Click
    Process.Start("notepad.exe", Application.StartupPath & "\IsgecMC.config")
  End Sub
End Class


Imports System.Data
Imports System.Data.SqlClient
Imports System.Net.Mail
Imports System.Net.Mail.SmtpClient
Public Class NewAttendanceRules
  Public Shared FinYear As String = ""
  Public Shared Property MayStopProcess As Boolean = False
  Public Shared Property smtpServer As String = ""
  Public Shared Property fromEmail As String = ""
  Public Shared Property fromID As String = ""
  Public Shared Property fromPW As String = ""
  Public Shared Property cc As String = ""
  Public Shared Property ConnectionString() As String
    Get
      Return SIS.SYS.SQLDatabase.DBCommon.conString
    End Get
    Set(value As String)
      SIS.SYS.SQLDatabase.DBCommon.conString = value
    End Set
  End Property
  Private Shared Function ActivePunchConfig(ByVal ProcessingDate As DateTime, ByVal OfficeID As String) As Integer
    Dim _Result As Integer = 0
    _Result = 6
    If ProcessingDate > Convert.ToDateTime("06/09/2015") Then
      _Result = 7
      If OfficeID = "4" Or OfficeID = "5" Or OfficeID = "6" Or OfficeID = "7" Or OfficeID = "8" Then
        _Result = 6
        If ProcessingDate >= Convert.ToDateTime("01/03/2018") Then
          If OfficeID = "5" Then
            _Result = 7
          End If
        End If
      End If
    End If
    If ProcessingDate >= Convert.ToDateTime("01/01/2021") Then
      If OfficeID = "4" Or OfficeID = "6" Or OfficeID = "7" Or OfficeID = "8" Then
        _Result = 13
      Else
        _Result = 12
      End If
    End If
    Return _Result
  End Function
  Private Shared ci As System.Globalization.CultureInfo = New System.Globalization.CultureInfo("en-GB", True)

  Public Shared Function GetOfficeID(ByVal OfficeID As Integer) As Integer
    Dim mRet As Integer = 1
    Select Case OfficeID
      Case 4
        mRet = 4
      Case 5
        mRet = 5
      Case 6
        mRet = 6
      Case Else
        mRet = 1
    End Select
    Return mRet
  End Function
  Private Shared Function GetEMailID(ByVal LoginID As String) As String
    Dim _Result As String = ""
    Using Con As SqlConnection = New SqlConnection(SIS.SYS.SQLDatabase.DBCommon.GetConnectionString)
      Using Cmd As SqlCommand = Con.CreateCommand()
        Dim mSql As String = "SELECT ISNULL(EMailID,'') FROM [aspnet_Users] WHERE UserName = '" & LoginID & "'"
        Cmd.CommandType = System.Data.CommandType.Text
        Cmd.CommandText = mSql
        Con.Open()
        _Result = Cmd.ExecuteScalar()
      End Using
    End Using
    Return _Result
  End Function

  Private Shared Sub PostAdvanceApplication(ByVal _atnd As SIS.ATN.atnNewAttendance)
    Dim _Lgrs As List(Of SIS.ATN.atnLeaveLedger) = SIS.ATN.atnLeaveLedger.GetByApplDetailID(_atnd.AttenID)
    'There is no Change, POST the application line
    If _atnd.PunchValue = 0 Then
      _atnd.Posted = True
      _atnd.ApplStatusID = 6
      _atnd.FinalValue = 1
      SIS.ATN.atnNewAttendance.Update(_atnd)
      For Each _lgr As SIS.ATN.atnLeaveLedger In _Lgrs
        _lgr.Days = -1 * _lgr.InProcessDays
        _lgr.InProcessDays = 0
        _lgr.TranDate = Now
        _lgr.FinYear = FinYear
        SIS.ATN.atnLeaveLedger.Update(_lgr)
      Next
    End If
    'There is Full Change, Employee Present
    If _atnd.PunchValue = 1 Then
      Dim _aplHeaderID As Integer = _atnd.ApplHeaderID
      With _atnd
        .Applied = False
        .AdvanceApplication = False
        .ApplHeaderID = ""
        .Applied1LeaveTypeID = ""
        .Applied2LeaveTypeID = ""
        .AppliedValue = ""
        .ApplStatusID = ""
        .Posted = False
        .Posted1LeaveTypeID = ""
        .Posted2LeaveTypeID = ""
        .FinalValue = 1
        .FinYear = FinYear
      End With
      SIS.ATN.atnNewAttendance.Update(_atnd)
      For Each _lgr As SIS.ATN.atnLeaveLedger In _Lgrs
        SIS.ATN.atnLeaveLedger.Delete(_lgr)
      Next
      'Delete the Header, If there is no more line in Application
      Dim _tmps As List(Of SIS.ATN.atnNewAttendance) = SIS.ATN.atnNewAttendance.GetAttendanceByApplHeaderID(_aplHeaderID)
      If _tmps.Count = 0 Then
        Dim _tmp As SIS.ATN.atnApplHeader = SIS.ATN.atnApplHeader.GetByID(_aplHeaderID)
        SIS.ATN.atnApplHeader.Delete(_tmp)
      End If
    End If
    'There is Partial Change
    If _atnd.PunchValue = 0.5 Then
      If _Lgrs.Count > 1 Then
        For Each _lgr As SIS.ATN.atnLeaveLedger In _Lgrs
          If _atnd.PunchStatusID = "AF" Then
            If _lgr.LeaveTypeID = _atnd.Posted2LeaveTypeID Then
              SIS.ATN.atnLeaveLedger.Delete(_lgr)
            End If
            If _lgr.LeaveTypeID = _atnd.Posted1LeaveTypeID Then
              _lgr.Days = -0.5
              _lgr.InProcessDays = 0
              _lgr.TranDate = Now
              _lgr.FinYear = FinYear
              SIS.ATN.atnLeaveLedger.Update(_lgr)
            End If
          Else ' Punch Status ID = AS,TS
            If _lgr.LeaveTypeID = _atnd.Posted1LeaveTypeID Then
              SIS.ATN.atnLeaveLedger.Delete(_lgr)
            End If
            If _lgr.LeaveTypeID = _atnd.Posted2LeaveTypeID Then
              _lgr.Days = -0.5
              _lgr.InProcessDays = 0
              _lgr.TranDate = Now
              _lgr.FinYear = FinYear
              SIS.ATN.atnLeaveLedger.Update(_lgr)
            End If
          End If
        Next
      Else 'Lgr.Count = 1
        For Each _lgr As SIS.ATN.atnLeaveLedger In _Lgrs
          _lgr.Days = -0.5
          _lgr.InProcessDays = 0
          _lgr.TranDate = Now
          _lgr.FinYear = FinYear
          SIS.ATN.atnLeaveLedger.Update(_lgr)
        Next
      End If
      With _atnd
        If _atnd.PunchStatusID = "AF" Then
          .Applied2LeaveTypeID = ""
          .Posted2LeaveTypeID = ""
        Else  'Punch Status ID = AS,TS
          .Applied1LeaveTypeID = ""
          .Posted1LeaveTypeID = ""
        End If
        .AppliedValue = 0.5
        .Posted = True
        .ApplStatusID = 6
        .FinalValue = 1
        .FinYear = FinYear
      End With
      SIS.ATN.atnNewAttendance.Update(_atnd)
    End If
    'Update Header
    'If Not complete change and header record is there
    If _atnd.ApplHeaderID <> String.Empty Then
      Dim _hdr As SIS.ATN.atnApplHeader = SIS.ATN.atnApplHeader.GetByID(_atnd.ApplHeaderID)
      With _hdr
        Dim Found As Boolean = False
        Dim oAplDet As List(Of SIS.ATN.atnNewAttendance) = SIS.ATN.atnNewAttendance.GetAttendanceByApplHeaderID(_hdr.LeaveApplID)
        For Each _det As SIS.ATN.atnNewAttendance In oAplDet
          If Not _det.Posted Then
            Found = True
            Exit For
          End If
        Next
        If Found Then
          .ApplStatusID = 5
          .ExecutionState = 2
        Else
          .ApplStatusID = 6
          .ExecutionState = 3
        End If
        .PostingRemark = "Auto Posted"
        .PostedOn = Now
        .PostedBy = "0340"
      End With
      SIS.ATN.atnApplHeader.Update(_hdr)
    End If
  End Sub
  Private Shared Function GetQtr(dd As DateTime) As Integer
    Select Case dd.Month
      Case 1, 2, 3
        Return 1
      Case 4, 5, 6
        Return 2
      Case 7, 8, 9
        Return 3
      Case Else
        Return 4
    End Select
  End Function
  Private Shared Function GetProrate(doj As DateTime, val As Decimal) As Decimal
    Dim mRet As Decimal = 0.00
    Select Case GetQtr(doj)
      Case 1
        Dim TotDays As Integer = 31 + DateTime.DaysInMonth(doj.Year, 2) + 31
        Dim WrkDays As Integer = 1 + (Convert.ToDateTime("31/03/" & doj.Year).Subtract(doj)).Days
        Return LvRoundOf(WrkDays * val / TotDays)
      Case 2
        Dim TotDays As Integer = 91
        Dim WrkDays As Integer = 1 + (Convert.ToDateTime("30/06/" & doj.Year).Subtract(doj)).Days
        Return LvRoundOf(WrkDays * val / TotDays)
      Case 3
        Dim TotDays As Integer = 92
        Dim WrkDays As Integer = 1 + (Convert.ToDateTime("30/09/" & doj.Year).Subtract(doj)).Days
        Return LvRoundOf(WrkDays * val / TotDays)
      Case Else
        Dim TotDays As Integer = 92
        Dim WrkDays As Integer = 1 + (Convert.ToDateTime("31/12/" & doj.Year).Subtract(doj)).Days
        Return LvRoundOf(WrkDays * val / TotDays)
    End Select
  End Function
  Private Shared Function GetProrateMon(doj As DateTime, val As Decimal) As Decimal
    Dim mRet As Decimal = 0.00
    Dim TotDays As Integer = DateTime.DaysInMonth(doj.Year, doj.Month)
    Dim WrkDays As Integer = TotDays - doj.Day + 1
    Return LvRoundOf(WrkDays * val / TotDays)
  End Function
  Private Shared Function LvRoundOf(ByVal nVal As Single) As Single
    Dim iVal As Integer
    Dim fVal As Single
    iVal = Int(nVal)
    fVal = nVal - iVal
    If fVal >= 0.75 Then
      fVal = 1
    Else
      If fVal >= 0.5 Then
        fVal = 0.5
      Else
        If fVal >= 0.25 Then
          fVal = 0.5
        Else
          fVal = 0
        End If
      End If
    End If
    nVal = iVal + fVal
    Return nVal
  End Function

  Public Shared Function ActualProcess(ByVal DataDate As DateTime, ByVal oEmp As SIS.ATN.atnEmployees) As String
    Dim mErr As String = ""
    'ProcessingDate = DataDate
    If oEmp.C_DateOfJoining = String.Empty Then
      Return "Erroe: DOJ blank"
    End If
    '=>Delete Punch Records of Before Joining
    'Do not process records of Not Joined By date
    If DateDiff(DateInterval.Day, DataDate, Convert.ToDateTime(oEmp.C_DateOfJoining, ci)) > 0 Then
      Return "DOJ: " & oEmp.C_DateOfJoining
    End If
    If oEmp.C_DateOfReleaving <> String.Empty Then
      '=>Delete Punch Data of After Releaving
      'Do not process Records of Releaved Employees by Releaving Date
      If DateDiff(DateInterval.Day, Convert.ToDateTime(oEmp.C_DateOfReleaving, ci), DataDate) > 0 Then
        Return "DOL: " & oEmp.C_DateOfReleaving
      End If
    End If
    Dim CardNo As String = oEmp.CardNo
    'Leave Credit Logic
    If DataDate.Year >= 2021 Then
      Dim xGr As New SIS.ATN.atnBalanceTransfer
      With xGr
        .CardNo = oEmp.CardNo
        .TranType = "OPB"
        .SubTranType = "MC"
        .TranDate = DataDate
        .LeaveTypeID = ""
        .FinYear = FinYear
        .Remarks = ""
        .Days = 0
        .FinYear = FinYear
      End With

      If oEmp.C_OfficeID <> 6 Then
        If oEmp.Contractual Then
          'Contractual Office Employee
          'Total 16 days, 4 days Qtr Credit
          Dim crDays As Decimal = 0
          If DataDate >= Convert.ToDateTime(oEmp.C_DateOfJoining) Then
            Dim oLgr As SIS.ATN.atnBalanceTransfer = SIS.ATN.atnBalanceTransfer.GetOPBRecordForQuarter(oEmp.CardNo, Month(DataDate), FinYear)
            If GetQtr(oEmp.C_DateOfJoining) = GetQtr(DataDate) AndAlso Year(oEmp.C_DateOfJoining) = Year(DataDate) Then
              crDays = GetProrate(oEmp.C_DateOfJoining, 4)
            Else
              crDays = 4
            End If
            If oLgr Is Nothing Then
              oLgr = xGr
              oLgr.LeaveTypeID = "Z1"
              oLgr.Days = crDays
              oLgr.Remarks = "Qtr Credit " & crDays & " Z1 for " & GetQtr(DataDate)
              SIS.ATN.atnBalanceTransfer.Insert(oLgr)
            Else
              If oLgr.Days <> crDays Then
                oLgr.Days = crDays
                oLgr.TranDate = DataDate
                SIS.ATN.atnBalanceTransfer.MonthlyCreditUpdate(oLgr)
              End If
            End If
          End If
        Else
          If oEmp.CreditQuarterly Then
            If oEmp.C_DesignationID = "34" Or oEmp.C_DesignationID = "35" Or oEmp.C_DesignationID = "28" Or oEmp.PRKCategoryID = 19 Or oEmp.PRKCategoryID = 18 Then
              'Conditional Employee
              'Total 20 days, 5 days qtr Credit
              Dim crDays As Decimal = 0
              If DataDate >= Convert.ToDateTime(oEmp.C_DateOfJoining) Then
                Dim oLgr As SIS.ATN.atnBalanceTransfer = SIS.ATN.atnBalanceTransfer.GetOPBRecordForQuarter(oEmp.CardNo, Month(DataDate), FinYear)
                If GetQtr(oEmp.C_DateOfJoining) = GetQtr(DataDate) AndAlso Year(oEmp.C_DateOfJoining) = Year(DataDate) Then
                  crDays = GetProrate(oEmp.C_DateOfJoining, 5)
                Else
                  crDays = 5
                End If
                If oLgr Is Nothing Then
                  oLgr = xGr
                  oLgr.LeaveTypeID = "Z1"
                  oLgr.Days = crDays
                  oLgr.Remarks = "Qtr Credit " & crDays & " Z1 for " & GetQtr(DataDate)
                  SIS.ATN.atnBalanceTransfer.Insert(oLgr)
                Else
                  If oLgr.Days <> crDays Then
                    oLgr.Days = crDays
                    oLgr.TranDate = DataDate
                    SIS.ATN.atnBalanceTransfer.MonthlyCreditUpdate(oLgr)
                  End If
                End If
              End If
            Else
              'General Office Employee
              'Total 30 days, 7.5 days Qtr Credit
              Dim crDays As Decimal = 0
              If DataDate >= Convert.ToDateTime(oEmp.C_DateOfJoining) Then
                Dim oLgr As SIS.ATN.atnBalanceTransfer = SIS.ATN.atnBalanceTransfer.GetOPBRecordForQuarter(oEmp.CardNo, Month(DataDate), FinYear)
                If GetQtr(oEmp.C_DateOfJoining) = GetQtr(DataDate) AndAlso Year(oEmp.C_DateOfJoining) = Year(DataDate) Then
                  crDays = GetProrate(oEmp.C_DateOfJoining, 7.5)
                Else
                  crDays = 7.5
                End If
                If oLgr Is Nothing Then
                  oLgr = xGr
                  oLgr.LeaveTypeID = "Z1"
                  oLgr.Days = crDays
                  oLgr.Remarks = "Qtr Credit " & crDays & " Z1 for " & GetQtr(DataDate)
                  SIS.ATN.atnBalanceTransfer.Insert(oLgr)
                Else
                  If oLgr.Days <> crDays Then
                    oLgr.Days = crDays
                    oLgr.TranDate = DataDate
                    SIS.ATN.atnBalanceTransfer.MonthlyCreditUpdate(oLgr)
                  End If
                End If
              End If
            End If
          End If
        End If
      Else
        If oEmp.Contractual Then
          'Contractual Site Employee
          'Total 24 days, 2 days Monthly Credit
          Dim crDays As Decimal = 0
          If DataDate >= Convert.ToDateTime(oEmp.C_DateOfJoining) Then
            Dim oLgr As SIS.ATN.atnBalanceTransfer = SIS.ATN.atnBalanceTransfer.GetOPBRecordForMonth(oEmp.CardNo, Month(DataDate), FinYear)
            If Month(oEmp.C_DateOfJoining) = Month(DataDate) AndAlso Year(oEmp.C_DateOfJoining) = Year(DataDate) Then
              crDays = GetProrateMon(oEmp.C_DateOfJoining, 2)
            Else
              crDays = 2
            End If
            If oLgr Is Nothing Then
              oLgr = xGr
              oLgr.LeaveTypeID = "CL"
              oLgr.Days = crDays
              oLgr.Remarks = "Monthly Credit " & crDays & " CL for " & MonthName(DataDate.Month, True)
              SIS.ATN.atnBalanceTransfer.Insert(oLgr)
            Else
              If oLgr.Days <> crDays Then
                oLgr.Days = crDays
                oLgr.TranDate = DataDate
                SIS.ATN.atnBalanceTransfer.MonthlyCreditUpdate(oLgr)
              End If
            End If
          End If
        Else
          'Old Rule CL,SL, PL(Next Year) =>Nothing to do
        End If
      End If
    Else
      'For Contractual Employee
      If oEmp.Contractual Then
        'Check Credited Leave record for the Month is created or not
        Dim oLgr As SIS.ATN.atnBalanceTransfer = SIS.ATN.atnBalanceTransfer.GetOPBRecordForMonth(oEmp.CardNo, Month(DataDate), FinYear)
        'Credit 2 CL for Month
        'On prorate with DOJ
        Dim crDays As Single = 0
        If Month(oEmp.C_DateOfJoining) = Month(DataDate) And Year(oEmp.C_DateOfJoining) = Year(DataDate) Then
          Select Case Convert.ToDateTime(oEmp.C_DateOfJoining).Day
            Case Is >= 24
              crDays = 2
            Case Is >= 17
              crDays = 1.5
            Case Is >= 12
              crDays = 1
            Case Is >= 4
              crDays = 0.5
          End Select
        Else
          If Convert.ToDateTime(oEmp.C_DateOfJoining, ci) < Convert.ToDateTime(DataDate, ci) Then
            crDays = 2
          End If
        End If
        If oLgr Is Nothing Then
          oLgr = New SIS.ATN.atnBalanceTransfer
          With oLgr
            .CardNo = oEmp.CardNo
            .TranType = "OPB"
            .SubTranType = "MC"
            .TranDate = DataDate
            .LeaveTypeID = "CL"
            .FinYear = FinYear
            .Remarks = "Monthly Credit " & crDays & " CL for " & MonthName(DataDate.Month, True)
            .Days = crDays
            .FinYear = FinYear
          End With
          SIS.ATN.atnBalanceTransfer.Insert(oLgr)
        Else
          With oLgr
            .CardNo = oEmp.CardNo
            .TranType = "OPB"
            .SubTranType = "MC"
            .TranDate = DataDate
            .LeaveTypeID = "CL"
            .FinYear = FinYear
            .Remarks = "Monthly Credit " & crDays & " CL for " & MonthName(DataDate.Month, True)
            .Days = crDays
            .FinYear = FinYear
          End With
          SIS.ATN.atnBalanceTransfer.MonthlyCreditUpdate(oLgr)
        End If
      End If
      'End For Contractual Employee

    End If



    Dim OfficeID As Integer = GetOfficeID(oEmp.C_OfficeID)
    Dim oHld As SIS.ATN.atnHolidays = SIS.ATN.atnHolidays.GetByHoliday(DataDate, OfficeID, FinYear)

    Dim oAtnd As SIS.ATN.atnNewAttendance = SIS.ATN.atnNewAttendance.GetAttendanceByCardNoDate(CardNo, DataDate)
    If oAtnd Is Nothing Then
      oAtnd = New SIS.ATN.atnNewAttendance
      With oAtnd
        .CardNo = CardNo
        .AttenDate = DataDate
        .Punch1Time = 0
        .Punch2Time = 0
        .PunchStatusID = "AD"
        .PunchValue = 0
        .FinalValue = 0
        .NeedsRegularization = True
        .HoliDay = IIf(oHld Is Nothing, False, True)
        .FinYear = FinYear
      End With
      oAtnd.AttenID = SIS.ATN.atnNewAttendance.Insert(oAtnd)
    Else
      oAtnd.HoliDay = IIf(oHld Is Nothing, False, True)
      SIS.ATN.atnNewAttendance.Update(oAtnd)
    End If

    'Update Punch Time and Status by New Values from Raw Punch
    '1. Only when user has not applied
    '2. User has applied but in advance
    '   2.1. This record is not posted
    '3. Attendance Record is not Mannually Edited
    '
    If Not oAtnd.MannuallyCorrected Then
      If Not oAtnd.Applied Then
        mErr = UpdateAttendanceFromRawData(oAtnd, OfficeID)
      Else
        If oAtnd.AdvanceApplication Then
          If Not oAtnd.Posted Then
            'Check If the Advance Application is Approved or not
            'If not Approved, Delete It,
            'If not under posting, i.e. not approved
            If oAtnd.ApplStatusID = 5 Then  'Under Posting
              mErr = UpdateAttendanceFromRawData(oAtnd, OfficeID)
              PostAdvanceApplication(oAtnd)
            Else
              mErr = UpdateAttendanceFromRawData(oAtnd, OfficeID)
              'Delete this Application Line, or complete Application
              'This Delete is Stopped in new system
              '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
              'DeleteAdvanceApplication(oAtnd)
              '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
            End If
          End If
        End If
      End If
    Else ' Mannually Corrected
      'Do not update card punch data from file
      'for this just delete the import line in coped logic
      If oAtnd.Applied Then
        If oAtnd.AdvanceApplication Then
          If Not oAtnd.Posted Then
            'Check If the Advance Application is Approved or not
            'If not Approved, Delete It,
            'If not under posting, i.e. not approved
            If oAtnd.ApplStatusID = 5 Then  'Under Posting
              PostAdvanceApplication(oAtnd)
            Else
              'Delete this Application Line, or complete Application
              'This Delete is Stopped in new system
              '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
              'DeleteAdvanceApplication(oAtnd)
              '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
            End If
          End If
        End If
      End If
    End If
    '========================================
    'In All Case
    If oEmp.C_OfficeID = 6 And Not oEmp.Contractual Then UpdateInterweavingHolidays(oAtnd.AttenID)
    '========================================
    Return oAtnd.PunchStatusID & ", " & oAtnd.Punch1Time & ", " & oAtnd.Punch2Time & "," & mErr
  End Function
  Private Shared Function IsValidPunchLocation(ByVal OfficeID As Integer, ByVal PunchLocation As Integer) As Boolean
    Dim mRet As Boolean = False
    If OfficeID = PunchLocation Then Return True
    'Site Employee can Punch in any Office
    If OfficeID = 6 Then Return True
    Using Con As SqlConnection = New SqlConnection(SIS.SYS.SQLDatabase.DBCommon.GetConnectionString())
      Using Cmd As SqlCommand = Con.CreateCommand()
        Cmd.CommandType = CommandType.Text
        Cmd.CommandText = "SELECT * FROM ATN_PunchLocationsByOffice WHERE OfficeID = " & OfficeID & " AND AllowedPunchLocation = " & PunchLocation
        Con.Open()
        Dim Reader As SqlDataReader = Cmd.ExecuteReader()
        If Reader.Read() Then
          mRet = True
        End If
        Reader.Close()
      End Using
    End Using
    Return mRet
  End Function
  Private Shared Function UpdateAttendanceFromRawData(ByVal oAtnd As SIS.ATN.atnNewAttendance, ByVal OfficeID As Integer) As String
    Dim mErr As String = ""
    Try
      Dim CardNo As String = oAtnd.CardNo
      Dim DataDate As DateTime = Convert.ToDateTime(oAtnd.AttenDate, ci)

      Dim AtndWasTS As Boolean = IIf(oAtnd.PunchStatusID = "TS", True, False)
      Dim AtndWasFP As Boolean = IIf(oAtnd.ConfigStatus = "FP" Or oAtnd.ConfigStatus = "LF", True, False)

      'Get Raw Punch
      Dim oRaw As SIS.ATN.atnRawPunch = SIS.ATN.atnRawPunch.GetRawPunchByCardNoDate(CardNo, DataDate)
      'Punch Required
      Dim oPunchReq As SIS.ATN.atnPunchRequired = SIS.ATN.atnPunchRequired.GetPunchRequiredByCardNo(CardNo)
      Dim UseException As Boolean = False
      If Not oPunchReq Is Nothing Then
        UseException = oPunchReq.RuleException
      End If
      'Active Configuration
      Dim oCnf As SIS.ATN.atnPunchConfig = SIS.ATN.atnPunchConfig.GetByID(ActivePunchConfig(DataDate, OfficeID))
      'Get Calender
      Dim oHld As SIS.ATN.atnHolidays = SIS.ATN.atnHolidays.GetByHoliday(DataDate, OfficeID, FinYear)


      If Not oAtnd.AdvanceApplication Then
        '1.
        If Not oHld Is Nothing Then
          With oAtnd
            .Punch1Time = 0
            .Punch2Time = 0
            .PunchStatusID = oHld.PunchStatusID
            .PunchValue = 1
            .FinalValue = 1
            .NeedsRegularization = False
            .FirstPunchMachine = ""
            .SecondPunchMachine = ""
            .FinYear = FinYear
          End With
          SIS.ATN.atnNewAttendance.Update(oAtnd)
          Return "" '**********
        End If
        '2. WFH
        Dim oWFH As SIS.ATN.WFHRooster = SIS.ATN.WFHRooster.GetByCardDate(oAtnd.CardNo, oAtnd.AttenDate)
        If oWFH IsNot Nothing Then
          If oWFH.WFHFullDay Then
            With oAtnd
              .Punch1Time = 0
              .Punch2Time = 0
              .PunchStatusID = "PR"
              .PunchValue = 1
              .FinalValue = 1
              .NeedsRegularization = False
              .FirstPunchMachine = ""
              .SecondPunchMachine = ""
              .FinYear = FinYear
              .ApplStatusID = 10
            End With
            SIS.ATN.atnNewAttendance.Update(oAtnd)
            Return ""  '**********
          End If
        Else
          oWFH = New SIS.ATN.WFHRooster
        End If
        If Not oWFH.WFHFullDay And oAtnd.ApplStatusID = "10" And oAtnd.Applied = False Then
          With oAtnd
            .ApplStatusID = ""
            .Punch1Time = 0
            .Punch2Time = 0
            .PunchStatusID = "AD"
            .PunchValue = 0
            .FinalValue = 0
            .NeedsRegularization = True
          End With
          SIS.ATN.atnNewAttendance.Update(oAtnd)
        End If
        '3.
        If Not oPunchReq Is Nothing Then
          If oPunchReq.NoPunch Then
            With oAtnd
              .Punch1Time = oCnf.STD1Time
              .Punch2Time = oCnf.STD2Time
              .PunchStatusID = "PR"
              .PunchValue = 1
              .FinalValue = 1
              .NeedsRegularization = False
              .FirstPunchMachine = ""
              .SecondPunchMachine = ""
              .FinYear = FinYear
            End With
            SIS.ATN.atnNewAttendance.Update(oAtnd)
            Return ""  '**********
          End If
        End If
        '4.
        If oRaw Is Nothing Then
          With oAtnd
            If OfficeID = 6 Then
              'For Site Employees
              '==========Modified 31st Dec 2017
              '.Punch1Time = 9
              '.Punch2Time = 17.45
              '.PunchStatusID = "PR"
              '.PunchValue = 1
              '.FinalValue = 0
              '.NeedsRegularization = False
              '=======New Values 31st Dec 2017====
              .Punch1Time = 0
              .Punch2Time = 0
              .PunchStatusID = "AD"
              .PunchValue = 0
              .FinalValue = 0
              .NeedsRegularization = True
              '=========End New Value=========
              .SiteAttendance = True
              .FirstPunchMachine = ""
              .SecondPunchMachine = ""
              .FinYear = FinYear
            Else
              .Punch1Time = 0
              .Punch2Time = 0
              .PunchStatusID = "AD"
              .PunchValue = 0
              .FinalValue = 0
              .NeedsRegularization = True
              .FirstPunchMachine = ""
              .SecondPunchMachine = ""
              .FinYear = FinYear
            End If
          End With
          SIS.ATN.atnNewAttendance.Update(oAtnd)
          Return ""  '**********
        End If
      End If 'End NOT Advance Application

      If Not oRaw Is Nothing Then
        '----Location Wise Card Punch Restriction of First Punch----
        '----Started from 19th March 2012----
        Dim MachineID As Integer = Convert.ToUInt32(oRaw.FirstPunchMachine.Split("-".ToCharArray)(0))
        '====================================
        ' Read Actual Office ID from Employee Table, Not the parameter passed to this function,
        ' Which is converted office ID for Holidays List
        Dim tEmp As SIS.ATN.atnEmployees = SIS.ATN.atnEmployees.GetByID(CardNo)
        '====================================
        'Location wise punch checking validation
        If Not oPunchReq Is Nothing Then
          If oPunchReq.AllLocation Then
            oAtnd.Punch1Time = oRaw.Punch1Time
            oAtnd.Punch2Time = oRaw.Punch2Time
          Else
            If Not IsValidPunchLocation(IIf(oAtnd.OfficeID = "", tEmp.C_OfficeID, oAtnd.OfficeID), MachineID) Then
              oAtnd.Punch1Time = oRaw.Punch2Time
              oAtnd.Punch2Time = 0.0
            Else
              oAtnd.Punch1Time = oRaw.Punch1Time
              oAtnd.Punch2Time = oRaw.Punch2Time
            End If
          End If
        Else
          If Not IsValidPunchLocation(IIf(oAtnd.OfficeID = "", tEmp.C_OfficeID, oAtnd.OfficeID), MachineID) Then
            oAtnd.Punch1Time = oRaw.Punch2Time
            oAtnd.Punch2Time = 0.0
          Else
            oAtnd.Punch1Time = oRaw.Punch1Time
            oAtnd.Punch2Time = oRaw.Punch2Time
          End If
        End If
        'Location Wise
        '----End Of Locationwise Card Punch Checking----
        oAtnd.FirstPunchMachine = oRaw.FirstPunchMachine
        oAtnd.SecondPunchMachine = oRaw.SecondPunchMachine
        SIS.ATN.atnNewAttendance.Update(oAtnd)
      End If
      'This is again if not adv applied
      If Not oAtnd.AdvanceApplication Then
        If Not oPunchReq Is Nothing Then
          If oPunchReq.OnePunch Then
            If Convert.ToDecimal(oAtnd.Punch1Time) > 0 Or Convert.ToDecimal(oAtnd.Punch2Time) > 0 Then
              With oAtnd
                .Punch1Time = oCnf.STD1Time
                .Punch2Time = oCnf.STD2Time
                .PunchStatusID = "PR"
                .PunchValue = 1
                .FinalValue = 1
                .NeedsRegularization = False
                .FirstPunchMachine = ""
                .SecondPunchMachine = ""
                .FinYear = FinYear
              End With
              SIS.ATN.atnNewAttendance.Update(oAtnd)
              Return ""  '**********
            End If
          End If
        End If
      End If
      Dim oConfs As List(Of SIS.ATN.atnPunchConfigDetails) = SIS.ATN.atnPunchConfigDetails.GetByConfigID(oCnf.RecordID, "")
      For Each cnf As SIS.ATN.atnPunchConfigDetails In oConfs
        If Not UseException Then
          If cnf.ExceptionRule Then Continue For
        Else
          If Not cnf.ExceptionRule Then Continue For
        End If
        If Convert.ToDecimal(oAtnd.Punch1Time) >= Convert.ToDecimal(cnf.FPStartTime) And Convert.ToDecimal(oAtnd.Punch1Time) <= Convert.ToDecimal(cnf.FPEndTime) Then
          If Convert.ToDecimal(oAtnd.Punch2Time) >= Convert.ToDecimal(cnf.SPStartTime) And Convert.ToDecimal(oAtnd.Punch2Time) <= Convert.ToDecimal(cnf.SPEndTime) Then
            Dim Matched As Boolean = True
            If cnf.CalculateHours Then
              Dim HrsWorked As Double = 0
              If cnf.UseDefined Then
                HrsWorked = DiffTime(cnf.DefinedTime, oAtnd.Punch2Time)
              Else
                HrsWorked = DiffTime(oAtnd.Punch1Time, oAtnd.Punch2Time)
              End If
              Select Case cnf.HoursStatus
                Case "<"
                  If HrsWorked >= Convert.ToDecimal(cnf.HoursValue) Then
                    Matched = False
                  End If
                Case ">="
                  If HrsWorked < Convert.ToDecimal(cnf.HoursValue) Then
                    Matched = False
                  End If
              End Select
            End If
            If Matched Then
              oAtnd.ConfigID = cnf.ConfigID
              oAtnd.ConfigDetailID = cnf.SerialNo
              oAtnd.ConfigStatus = cnf.ConfigStatus
              If cnf.LimitedAllowed Then
                'first update oatnd then get count
                SIS.ATN.atnNewAttendance.Update(oAtnd)
                Dim LimitCount As Integer = SIS.ATN.atnNewAttendance.GetConfigCount(CardNo, DataDate, cnf.ConfigStatus, FinYear)
                Dim Availed As Integer = SIS.ATN.atnNewAttendance.GetFPAvailed(CardNo, DataDate, FinYear)
                If LimitCount <= Convert.ToInt32(cnf.LimitCount) Then
                  oAtnd.PunchStatusID = cnf.LimitPunchStatus
                Else
                  oAtnd.PunchStatusID = cnf.NormalPunchStatus
                End If
                'To Send Mail In Case of FP

                If cnf.ConfigStatus = "FP" Or cnf.ConfigStatus = "LF" Then
                  If Not AtndWasFP Then
                    Dim EMailID As String = GetEMailID(oAtnd.CardNo)
                    If EMailID <> "" Then
                      Try
                        Dim oClient As SmtpClient = New SmtpClient(smtpServer)
                        Dim cr As System.Net.ICredentialsByHost = New System.Net.NetworkCredential(fromID, fromPW)
                        oClient.Credentials = cr
                        Dim oMsg As New System.Net.Mail.MailMessage
                        Dim oEmp As SIS.ATN.atnEmployees = SIS.ATN.atnEmployees.GetByID(oAtnd.CardNo)
                        With oMsg
                          .From = New MailAddress(fromEmail)
                          .IsBodyHtml = True
                          .To.Add(EMailID)
                          .CC.Add(cc)
                          .Subject = "ISGEC Attendance System [FP]"
                          Dim mStr As String = "<table border=""0"" cellspacing=""0"" cellpadding=""0"">"
                          mStr = mStr & "<tr><td><b>Dear " & oEmp.EmployeeName & ",</b></td></tr>"
                          mStr = mStr & "<tr><td><b>Attendance Date :</b>" & oAtnd.AttenDate & "</td></tr>"
                          mStr = mStr & "<tr><td><b>Ist Punch Time :</b>" & oAtnd.Punch1Time & "</td></tr>"
                          mStr = mStr & "<tr><td><b>IInd Punch Time :</b>00.00</td></tr>"
                          mStr = mStr & "<tr><td>Your IInd card punch not found.</td></tr>"
                          'If Availed <= Convert.ToInt32(cnf.LimitCount) Then
                          '  mStr = mStr & "<tr><td>Clear your absent status using Regularize Forget Punch.</td></tr>"
                          '  mStr = mStr & "<tr><td>Your can Regularize FP upto <b>" & cnf.LimitCount & "</b> times. Allready used : <b>" & Availed & "</b></td></tr>"
                          'End If
                          mStr = mStr & "<tr><td>Please regularize your attendance to avoid deduction.</td></tr>"
                          mStr = mStr & "<tr><td>Thanx.</td></tr>"
                          mStr = mStr & "</table>"
                          .Body = mStr
                        End With
                        'oClient.Send(oMsg)

                      Catch ex As Exception
                        mErr = ex.Message
                      End Try
                    End If
                  End If
                End If
                'End Of FP Mail
              Else 'No Limit
                oAtnd.PunchStatusID = cnf.NormalPunchStatus
              End If
              'It Is Hard Coded Must Be Picked From ATN_PunchStatus
              ' Add One More Field in NeedsRegularization
              Select Case oAtnd.PunchStatusID
                Case "PR"
                  oAtnd.PunchValue = 1
                  oAtnd.FinalValue = 1
                  oAtnd.NeedsRegularization = False
                Case "AF", "AS"
                  oAtnd.PunchValue = 0.5
                  oAtnd.FinalValue = 0.5
                  oAtnd.NeedsRegularization = True
                Case "AD"
                  oAtnd.PunchValue = 0
                  oAtnd.FinalValue = 0
                  oAtnd.NeedsRegularization = True
                Case "TS"
                  oAtnd.PunchValue = 0.5
                  oAtnd.FinalValue = 0.5     'Auto Leave Not Posted
                  oAtnd.NeedsRegularization = False
                  oAtnd.TSStatus = "TS"
              End Select
              If oAtnd.AdvanceApplication Then
                oAtnd.AppliedValue = 1 - oAtnd.FinalValue
              End If
              SIS.ATN.atnNewAttendance.Update(oAtnd)

              If oAtnd.PunchStatusID <> "TS" Then
              Else
                If Not AtndWasTS Then
                  Dim EMailID As String = GetEMailID(oAtnd.CardNo)
                  If EMailID <> "" Then
                    Try
                      Dim oClient As SmtpClient = New SmtpClient(smtpServer)
                      Dim cr As System.Net.ICredentialsByHost = New System.Net.NetworkCredential(fromID, fromPW)
                      oClient.Credentials = cr
                      Dim oMsg As New System.Net.Mail.MailMessage
                      Dim oEmp As SIS.ATN.atnEmployees = SIS.ATN.atnEmployees.GetByID(oAtnd.CardNo)
                      With oMsg
                        .From = New MailAddress(fromEmail)
                        .IsBodyHtml = True
                        .To.Add(EMailID)
                        .CC.Add(cc)
                        .Subject = "ISGEC Attendance System [TS]"
                        Dim mStr As String = "<table border=""0"" cellspacing=""0"" cellpadding=""0"">"
                        mStr = mStr & "<tr><td><b>Dear " & oEmp.EmployeeName & ",</b></td></tr>"
                        mStr = mStr & "<tr><td><b>Attendance Date :</b>" & oAtnd.AttenDate & "</td></tr>"
                        mStr = mStr & "<tr><td><b>Ist Punch Time :</b>" & oAtnd.Punch1Time & "</td></tr>"
                        mStr = mStr & "<tr><td><b>IInd Punch Time :</b>" & oAtnd.Punch2Time & "</td></tr>"
                        mStr = mStr & "<tr><td>Required working hours/day is short.</td></tr>"
                        mStr = mStr & "<tr><td>For further clarification in this regard, you may contact to Time Office.</td></tr>"
                        mStr = mStr & "<tr><td>Thanx.</td></tr>"
                        mStr = mStr & "</table>"
                        .Body = mStr
                      End With

                      'oClient.Send(oMsg)

                    Catch ex As Exception
                      mErr = ex.Message
                    End Try
                  End If
                End If
              End If
              'When Matched Exit Loop
              Exit For
            End If 'End of Matched
          End If
        End If
      Next
    Catch ex As Exception
      Throw ex
    End Try
    Return mErr
  End Function

  Private Shared Sub RevertHLD(ByVal _sts As AtnSts)
    Dim oHld As SIS.ATN.atnHolidays = Nothing
    oHld = SIS.ATN.atnHolidays.GetByHoliday(_sts.AttenDate, _sts.OfficeID, FinYear)
    If Not oHld Is Nothing Then
      Do While Not oHld Is Nothing
        Dim oTmp As SIS.ATN.atnProcessedPunch = SIS.ATN.atnProcessedPunch.GetProcessedPunchByCardNoDate(_sts.CardNo, _sts.AttenDate)
        If Not oTmp Is Nothing Then
          If Not oTmp.Applied Then
            If oTmp.FinalValue = 0 Then
              With oTmp
                .Punch1Time = 0
                .Punch2Time = 0
                .PunchStatusID = oHld.PunchStatusID
                .PunchValue = 1
                .FinalValue = 1
                .NeedsRegularization = False
                .FinYear = FinYear
              End With
              SIS.ATN.atnProcessedPunch.Update(oTmp)
            End If
          End If
        End If
        _sts = GetPSts(_sts)
        If _sts Is Nothing Then
          Exit Do
        End If
        oHld = SIS.ATN.atnHolidays.GetByHoliday(_sts.AttenDate, _sts.OfficeID, FinYear)
      Loop
    End If
  End Sub
  Private Shared Sub AbsentHLD(ByVal _sts As AtnSts)
    Dim oHld As SIS.ATN.atnHolidays = Nothing
    oHld = SIS.ATN.atnHolidays.GetByHoliday(_sts.AttenDate, _sts.OfficeID, FinYear)
    If Not oHld Is Nothing Then
      Do While Not oHld Is Nothing
        Dim oTmp As SIS.ATN.atnProcessedPunch = SIS.ATN.atnProcessedPunch.GetProcessedPunchByCardNoDate(_sts.CardNo, _sts.AttenDate)
        If Not oTmp Is Nothing Then
          If Not oTmp.Applied Then
            If oTmp.FinalValue = 1 Then
              With oTmp
                .PunchStatusID = "AD"
                .PunchValue = 0
                .FinalValue = 0
                .NeedsRegularization = True
                .FinYear = FinYear
              End With
              SIS.ATN.atnProcessedPunch.Update(oTmp)
            End If
          End If
        End If
        _sts = GetPSts(_sts)
        If _sts Is Nothing Then
          Exit Do
        End If
        oHld = SIS.ATN.atnHolidays.GetByHoliday(_sts.AttenDate, _sts.OfficeID, FinYear)
      Loop
    End If
  End Sub
  Private Shared Function GetPSts(ByVal cSts As AtnSts) As AtnSts
    Dim _atnd As SIS.ATN.atnNewAttendance = SIS.ATN.atnNewAttendance.GetAttendanceByCardNoDate(cSts.CardNo, Convert.ToDateTime(cSts.AttenDate, ci).AddDays(-1))
    If Not _atnd Is Nothing Then
      Return New AtnSts(_atnd, cSts.OfficeID, cSts.CategoryID, cSts.Contractual)
    End If
    Return New AtnSts
  End Function
  Private Shared Function DiffTime(ByVal T1 As Double, ByVal T2 As Double) As Double
    Dim d1 As DateTime = DateAndTime.TimeSerial(Math.Floor(T1), (T1 - Math.Floor(T1)) * 100, 0)
    Dim d2 As DateTime = DateAndTime.TimeSerial(Math.Floor(T2), (T2 - Math.Floor(T2)) * 100, 0)
    Dim t3 As Integer = DateDiff(DateInterval.Minute, d1, d2)
    Dim t4 As Integer = Math.Floor(t3 / 60)
    Dim t5 As Integer = t3 Mod 60
    Dim t6 As Double = Convert.ToInt32(t4) + (t5 / 100)
    Return t6
  End Function

  Private Shared Sub UpdateInterweavingHolidays(ByVal AttenID As Integer)

    Dim tSts As AtnSts = Nothing
    Dim pSts As AtnSts = Nothing
    Dim hSts As AtnSts = Nothing

    Try


      Dim cSts As AtnSts = New AtnSts(SIS.ATN.atnNewAttendance.GetByID(AttenID))

      Select Case cSts.Status
        Case "PR"
          pSts = GetPSts(cSts)
          If pSts.Status = "PR" Then
            'do nothing
          ElseIf pSts.Status = "HD" Then
            'do nothing
          ElseIf pSts.Status = "AD" Then
            If pSts.Holiday Then
              RevertHLD(pSts)
            End If
          End If
        Case "HD"
          pSts = GetPSts(cSts)
          If pSts.Status = "PR" Then
            'allready HD
            'do nothing
          ElseIf pSts.Status = "HD" Then
            'allready HD
            'do nothing
          ElseIf pSts.Status = "AD" Then
            If pSts.LeaveType <> "CL" Then 'Any Leave Type or Not Applied or OD
              AbsentHLD(cSts)
            ElseIf pSts.LeaveType = "CL" Then
              tSts = GetPSts(pSts)
              Do While Not tSts Is Nothing
                If tSts.LeaveType <> "CL" Then
                  Exit Do
                End If
                tSts = GetPSts(tSts)
              Loop
              If Not tSts Is Nothing Then
                If tSts.Status = "AD" Then
                  If tSts.LeaveType <> "OD" Then  'not applied or any other Leavetype
                    AbsentHLD(cSts)
                  End If
                End If
              End If
            End If
          End If
        Case "AD"
          pSts = GetPSts(cSts)
          If pSts.Status = "PR" Then
            'do nothing
          ElseIf pSts.Status = "HD" Then
            'check to convert HD -> AD
            'in both cases Applied/Not Applied
            If cSts.LeaveType = "OD" Then   '$$$$Or cSts.LeaveType = "CL"
              'do nothing
            Else 'not applied or any other leave type
              tSts = GetPSts(pSts)
              Do While Not tSts Is Nothing
                If tSts.Status <> "HD" Then
                  Exit Do
                End If
                tSts = GetPSts(tSts)
              Loop
              If Not tSts Is Nothing Then
                If tSts.Status = "AD" Then
                  If tSts.LeaveType <> "OD" And tSts.LeaveType <> "CL" Then
                    AbsentHLD(pSts)
                  ElseIf tSts.LeaveType = "CL" Then
                    hSts = GetPSts(tSts)
                    Do While Not hSts Is Nothing
                      If hSts.LeaveType <> "CL" Then
                        Exit Do
                      End If
                      hSts = GetPSts(hSts)
                    Loop
                    If Not hSts Is Nothing Then
                      If hSts.Status = "AD" Then
                        If hSts.LeaveType <> "OD" Then
                          AbsentHLD(pSts)
                        End If
                      End If
                    End If
                  End If
                End If
              End If
            End If
          ElseIf pSts.Status = "AD" Then
            If pSts.Holiday Then
              'check to convert AD -> HD
              'in both cases Applied/Not Applied
              If cSts.LeaveType = "OD" Or cSts.LeaveType = "CL" Then
                RevertHLD(pSts)
              Else ' Not Applied or any Other Leave Type, then check Other end of holiday
                tSts = GetPSts(pSts)
                Do While Not tSts Is Nothing
                  If Not tSts.Holiday Then
                    Exit Do
                  End If
                  tSts = GetPSts(tSts)
                Loop
                If Not tSts Is Nothing Then
                  If tSts.Status = "PR" Or tSts.Status = "HD" Then
                    RevertHLD(pSts)
                  ElseIf tSts.Status = "AD" Then
                    If tSts.LeaveType = "OD" Then
                      RevertHLD(pSts)
                    ElseIf tSts.LeaveType = "CL" Then
                      hSts = GetPSts(tSts)
                      Do While Not hSts Is Nothing
                        If hSts.LeaveType <> "CL" Then
                          Exit Do
                        End If
                        hSts = GetPSts(hSts)
                      Loop
                      If Not hSts Is Nothing Then
                        If hSts.Status = "PR" Or hSts.Status = "HD" Then
                          RevertHLD(pSts)
                        ElseIf hSts.Status = "AD" Then
                          If hSts.LeaveType = "OD" Then
                            RevertHLD(pSts)
                          End If
                        End If
                      Else
                        RevertHLD(pSts)
                      End If
                    End If
                  End If
                Else
                  RevertHLD(pSts)
                End If
              End If
            End If
          End If
      End Select
    Catch ex As Exception
      Throw ex
    End Try
  End Sub
  Public Shared Sub PostUnpostedAdavanceApplication(ByVal dt As DateTime, ByVal CardNo As String)
    Dim oAtnds As New List(Of SIS.ATN.atnNewAttendance)
    Using Con As SqlConnection = New SqlConnection(SIS.SYS.SQLDatabase.DBCommon.GetConnectionString())
      Using Cmd As SqlCommand = Con.CreateCommand()
        Cmd.CommandType = CommandType.Text
        Cmd.CommandText = "SELECT * FROM ATNv_UnpostedAdvanceApplication WHERE CardNo = '" & CardNo & "' AND  AttenDate <= CONVERT(DATETIME,'" & dt.ToString("dd/MM/yyyy") & "', 103) AND FinYear ='" & FinYear & "'"
        Con.Open()
        Dim Reader As SqlDataReader = Cmd.ExecuteReader()
        If Reader.Read() Then
          oAtnds.Add(New SIS.ATN.atnNewAttendance(Reader))
        End If
        Reader.Close()
      End Using
    End Using
    For Each oAtnd As SIS.ATN.atnNewAttendance In oAtnds
      If oAtnd.AdvanceApplication Then
        If Not oAtnd.Posted Then
          'Check If the Advance Application is Approved or not
          'If not Approved, Delete It,
          'If not under posting, i.e. not approved
          '====
          Dim oEmp As SIS.ATN.atnEmployees = SIS.ATN.atnEmployees.GetByID(oAtnd.CardNo)
          Dim ImportRawData As Boolean = True
          Dim OfficeID As Integer = GetOfficeID(oEmp.C_OfficeID)
          '====
          If oAtnd.ApplStatusID = 5 Then  'Under Posting
            If ImportRawData Then
              UpdateAttendanceFromRawData(oAtnd, OfficeID)
            End If
            PostAdvanceApplication(oAtnd)
            If oEmp.C_OfficeID = 6 And Not oEmp.Contractual Then UpdateInterweavingHolidays(oAtnd.AttenID)
          Else
            If ImportRawData Then
              UpdateAttendanceFromRawData(oAtnd, OfficeID)
            End If
            If oEmp.C_OfficeID = 6 And Not oEmp.Contractual Then UpdateInterweavingHolidays(oAtnd.AttenID)
            'Delete this Application Line, or complete Application
            'This Delete is Stopped in new system
            '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
            'DeleteAdvanceApplication(oAtnd)
            '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
            'UpdateInterweavingHolidays(oAtnd.AttenID)
          End If
        End If
      End If
    Next
  End Sub

  'Public Shared Function ProcessData(ByVal FromDate As String, ByVal ToDate As String) As String
  '  Dim fdt As DateTime = Convert.ToDateTime(FromDate, ci)
  '  Dim tdt As DateTime = Convert.ToDateTime(ToDate, ci)
  '  Dim xdt As DateTime = fdt
  '  'Without Employee Code
  '  Dim StartRow As Integer = 0
  '  Dim MaxRow As Integer = 10
  '  Dim oEmps As List(Of SIS.ATN.atnEmployees) = SIS.ATN.atnEmployees.SelectList(StartRow, MaxRow, "CardNo", False, "")
  '  Do While oEmps.Count > 0
  '    For Each oEmp As SIS.ATN.atnEmployees In oEmps
  '      '================================
  '      'If C_OfficeID is blank then skip
  '      '================================
  '      If oEmp.C_OfficeID = "" Then
  '        'ErrStr = xdt.ToString("dd/MM/yyyy") & ", " & oEmp.CardNo & ", " & "Not Verified, Current Location is Blank"
  '        Continue For
  '      End If
  '      '================================
  '      Try
  '        xdt = fdt
  '        While xdt <= tdt
  '          FinYear = Year(xdt)
  '          ActualProcess(xdt, oEmp)
  '          xdt = xdt.AddDays(1)
  '        End While
  '      Catch ex As Exception
  '        'ErrStr = xdt.ToString("dd/MM/yyyy") & ", " & oEmp.CardNo & ", " & ex.Message
  '      End Try
  '      If MayStopProcess Then
  '        Exit For
  '      End If
  '    Next
  '    If MayStopProcess Then
  '      Exit Do
  '    End If
  '    StartRow += MaxRow
  '    oEmps = SIS.ATN.atnEmployees.SelectList(StartRow, MaxRow, "CardNo", False, "")
  '  Loop
  '  '=====================
  '  'Process Advance Application
  '  Try
  '    If Not MayStopProcess Then
  '      PostUnpostedAdavanceApplication(tdt)
  '    End If
  '  Catch ex As Exception
  '    'ErrStr = ex.Message
  '  End Try
  '  '=====================
  '  Return ""
  'End Function

End Class
Public Class AtnSts
  Private _CardNo As String = ""
  Private _OfficeID As Integer = 1
  Private _AttenID As Integer
  Private _Status As String = ""
  Private _LeaveType As String = ""
  Private _AttenDate As String = ""
  Private _Holiday As Boolean = False
  Private _CategoryID As String = ""
  Private _Contractual As Boolean = False
  Public Property Contractual() As Boolean
    Get
      Return _Contractual
    End Get
    Set(ByVal value As Boolean)
      _Contractual = value
    End Set
  End Property
  Public Property CategoryID() As String
    Get
      Return _CategoryID
    End Get
    Set(ByVal value As String)
      _CategoryID = value
    End Set
  End Property
  Public Property OfficeID() As Integer
    Get
      Return _OfficeID
    End Get
    Set(ByVal value As Integer)
      _OfficeID = value
    End Set
  End Property
  Public Property CardNo() As String
    Get
      Return _CardNo
    End Get
    Set(ByVal value As String)
      _CardNo = value
    End Set
  End Property
  Public Property AttenID() As Integer
    Get
      Return _AttenID
    End Get
    Set(ByVal value As Integer)
      _AttenID = value
    End Set
  End Property
  Public Property Holiday() As Boolean
    Get
      Return _Holiday
    End Get
    Set(ByVal value As Boolean)
      _Holiday = value
    End Set
  End Property
  Public Property AttenDate() As String
    Get
      Return _AttenDate
    End Get
    Set(ByVal value As String)
      _AttenDate = value
    End Set
  End Property
  Public Property LeaveType() As String
    Get
      Return _LeaveType
    End Get
    Set(ByVal value As String)
      _LeaveType = value
    End Set
  End Property
  Public Property Status() As String
    Get
      Return _Status
    End Get
    Set(ByVal value As String)
      _Status = value
    End Set
  End Property
  Public Sub New()
    'dummy
  End Sub
  Private Sub InitMe(ByVal pAtnd As SIS.ATN.atnNewAttendance, ByVal pOfficeID As Integer, ByVal pCategoryID As String, ByVal pContractual As Boolean)
    _AttenID = pAtnd.AttenID
    _AttenDate = pAtnd.AttenDate
    _Status = pAtnd.PunchStatusID
    _CardNo = pAtnd.CardNo
    _OfficeID = pOfficeID
    _CategoryID = pCategoryID
    _Contractual = pContractual
    _LeaveType = ""
    _Holiday = False
    Select Case pAtnd.PunchStatusID
      Case "AS", "AF", "TS", "PR", "NH"
        _Status = "PR"
      Case "WO", "HD"
        _Status = "HD"
        _Holiday = True
      Case "AD"
        If _CategoryID = "18" Or _CategoryID = "19" Or _Contractual Then
          _Status = "PR"
        ElseIf pAtnd.Applied And pAtnd.ApplStatusID = "6" Then
          If pAtnd.Applied1LeaveTypeID = "PL" Or pAtnd.Applied2LeaveTypeID = "PL" Then
            _LeaveType = "PL"
          End If
          If pAtnd.Applied1LeaveTypeID = "SL" Or pAtnd.Applied2LeaveTypeID = "SL" Then
            _LeaveType = "SL"
          End If
          If pAtnd.Applied1LeaveTypeID = "CL" Or pAtnd.Applied2LeaveTypeID = "CL" Then
            _LeaveType = "CL"
          End If
          If pAtnd.Applied1LeaveTypeID = "" And pAtnd.Applied2LeaveTypeID <> "" Then
            _LeaveType = pAtnd.Applied2LeaveTypeID
          End If
          If pAtnd.Applied1LeaveTypeID <> "" And pAtnd.Applied2LeaveTypeID = "" Then
            _LeaveType = pAtnd.Applied1LeaveTypeID
          End If
          If pAtnd.Applied1LeaveTypeID = "OD" Or pAtnd.Applied2LeaveTypeID = "OD" Then
            _LeaveType = "OD"
          End If
          If pAtnd.Applied1LeaveTypeID = "FP" Or pAtnd.Applied2LeaveTypeID = "FP" Then
            _LeaveType = "FP"
            _Status = "PR"
          End If
        End If
        Dim ci As System.Globalization.CultureInfo = New System.Globalization.CultureInfo("en-GB", True)
        Dim _hld As SIS.ATN.atnHolidays = SIS.ATN.atnHolidays.GetByHoliday(Convert.ToDateTime(pAtnd.AttenDate, ci), OfficeID, pAtnd.FinYear)
        If Not _hld Is Nothing Then
          _Holiday = True
        End If
    End Select
  End Sub
  Public Sub New(ByVal pAtnd As SIS.ATN.atnNewAttendance)
    Dim oEmp As SIS.ATN.atnEmployees = SIS.ATN.atnEmployees.GetByID(pAtnd.CardNo)
    InitMe(pAtnd, NewAttendanceRules.GetOfficeID(oEmp.C_OfficeID), GetCategoryID(pAtnd.CardNo), oEmp.Contractual)
  End Sub
  Public Sub New(ByVal pAtnd As SIS.ATN.atnNewAttendance, ByVal pOfficeID As Integer)
    Dim oEmp As SIS.ATN.atnEmployees = SIS.ATN.atnEmployees.GetByID(pAtnd.CardNo)
    InitMe(pAtnd, pOfficeID, GetCategoryID(pAtnd.CardNo), oEmp.Contractual)
  End Sub
  Public Sub New(ByVal pAtnd As SIS.ATN.atnNewAttendance, ByVal pOfficeID As Integer, ByVal pCategoryID As String)
    Dim oEmp As SIS.ATN.atnEmployees = SIS.ATN.atnEmployees.GetByID(pAtnd.CardNo)
    InitMe(pAtnd, pOfficeID, pCategoryID, oEmp.Contractual)
  End Sub
  Public Sub New(ByVal pAtnd As SIS.ATN.atnNewAttendance, ByVal pOfficeID As Integer, ByVal pCategoryID As String, ByVal pContractual As Boolean)
    InitMe(pAtnd, pOfficeID, pCategoryID, pContractual)
  End Sub
  Public Shared Function GetCategoryID(ByVal CardNo As String) As Integer
    Dim Results As Integer = 0
    Using Con As SqlConnection = New SqlConnection(SIS.SYS.SQLDatabase.DBCommon.GetConnectionString())
      Using Cmd As SqlCommand = Con.CreateCommand()
        Cmd.CommandType = CommandType.Text
        Cmd.CommandText = "SELECT CATEGORYID FROM PRK_EMPLOYEES WHERE CARDNO='" & CardNo & "'"
        Con.Open()
        Dim r As Object = Cmd.ExecuteScalar
        If Not r Is Nothing Then
          If Not r.ToString = String.Empty Then
            Results = Convert.ToInt32(r)
          End If
        End If
      End Using
    End Using
    Return Results
  End Function

End Class

Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Namespace SIS.ATN
  <DataObject()> _
  Partial Public Class atnNewAttendance
    Private Shared _RecordCount As Integer
    Private _AttenID As Int32
    Private _AttenDate As String
    Private _CardNo As String
    Private _Punch1Time As String
    Private _Punch2Time As String
    Private _PunchStatusID As String
    Private _Punch9Time As String
    Private _PunchValue As String
    Private _NeedsRegularization As Boolean
    Private _FinYear As String
    Private _Applied As Boolean
    Private _AppliedValue As String
    Private _Applied1LeaveTypeID As String
    Private _Applied2LeaveTypeID As String
    Private _Posted As Boolean
    Private _Posted1LeaveTypeID As String
    Private _Posted2LeaveTypeID As String
    Private _ApplHeaderID As String
    Private _ApplStatusID As String
    Private _FinalValue As String
    Private _AdvanceApplication As Boolean
    Private _MannuallyCorrected As Boolean
    Private _Destination As String
    Private _Purpose As String
    Private _ConfigID As String
    Private _ConfigDetailID As String
    Private _ConfigStatus As String
    Private _TSStatus As String
    Private _TSStatusBy As String
    Private _TSStatusOn As String
		Private _FirstPunchMachine As String = ""
		Private _SecondPunchMachine As String = ""
		Public Property SecondPunchMachine() As String
			Get
				Return _SecondPunchMachine
			End Get
			Set(ByVal value As String)
				If Convert.IsDBNull(value) Then
					_SecondPunchMachine = ""
				Else
					_SecondPunchMachine = value
				End If
			End Set
		End Property
		Public Property FirstPunchMachine() As String
			Get
				Return _FirstPunchMachine
			End Get
			Set(ByVal value As String)
				If Convert.IsDBNull(value) Then
					_FirstPunchMachine = ""
				Else
					_FirstPunchMachine = value
				End If
			End Set
		End Property
		Public Property AttenID() As Int32
			Get
				Return _AttenID
			End Get
			Set(ByVal value As Int32)
				_AttenID = value
			End Set
		End Property
    Public Property AttenDate() As String
      Get
        If Not _AttenDate = String.Empty Then
          Return Convert.ToDateTime(_AttenDate).ToString("dd/MM/yyyy")
        End If
        Return _AttenDate
      End Get
      Set(ByVal value As String)
			   _AttenDate = value
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
    Public Property Punch1Time() As String
      Get
        Return _Punch1Time
      End Get
      Set(ByVal value As String)
				 If Convert.IsDBNull(Value) Then
					 _Punch1Time = ""
				 Else
					 _Punch1Time = value
			   End If
      End Set
    End Property
    Public Property Punch2Time() As String
      Get
        Return _Punch2Time
      End Get
      Set(ByVal value As String)
				 If Convert.IsDBNull(Value) Then
					 _Punch2Time = ""
				 Else
					 _Punch2Time = value
			   End If
      End Set
    End Property
    Public Property PunchStatusID() As String
      Get
        Return _PunchStatusID
      End Get
      Set(ByVal value As String)
				 If Convert.IsDBNull(Value) Then
					 _PunchStatusID = ""
				 Else
					 _PunchStatusID = value
			   End If
      End Set
    End Property
    Public Property Punch9Time() As String
      Get
        Return _Punch9Time
      End Get
      Set(ByVal value As String)
				 If Convert.IsDBNull(Value) Then
					 _Punch9Time = ""
				 Else
					 _Punch9Time = value
			   End If
      End Set
    End Property
    Public Property PunchValue() As String
      Get
        Return _PunchValue
      End Get
      Set(ByVal value As String)
				 If Convert.IsDBNull(Value) Then
					 _PunchValue = ""
				 Else
					 _PunchValue = value
			   End If
      End Set
    End Property
    Public Property NeedsRegularization() As Boolean
      Get
        Return _NeedsRegularization
      End Get
      Set(ByVal value As Boolean)
        _NeedsRegularization = value
      End Set
    End Property
    Public Property FinYear() As String
      Get
        Return _FinYear
      End Get
      Set(ByVal value As String)
				 If Convert.IsDBNull(Value) Then
					 _FinYear = ""
				 Else
					 _FinYear = value
			   End If
      End Set
    End Property
    Public Property Applied() As Boolean
      Get
        Return _Applied
      End Get
      Set(ByVal value As Boolean)
        _Applied = value
      End Set
    End Property
    Public Property AppliedValue() As String
      Get
        Return _AppliedValue
      End Get
      Set(ByVal value As String)
				 If Convert.IsDBNull(Value) Then
					 _AppliedValue = ""
				 Else
					 _AppliedValue = value
			   End If
      End Set
    End Property
    Public Property Applied1LeaveTypeID() As String
      Get
        Return _Applied1LeaveTypeID
      End Get
      Set(ByVal value As String)
				 If Convert.IsDBNull(Value) Then
					 _Applied1LeaveTypeID = ""
				 Else
					 _Applied1LeaveTypeID = value
			   End If
      End Set
    End Property
    Public Property Applied2LeaveTypeID() As String
      Get
        Return _Applied2LeaveTypeID
      End Get
      Set(ByVal value As String)
				 If Convert.IsDBNull(Value) Then
					 _Applied2LeaveTypeID = ""
				 Else
					 _Applied2LeaveTypeID = value
			   End If
      End Set
    End Property
    Public Property Posted() As Boolean
      Get
        Return _Posted
      End Get
      Set(ByVal value As Boolean)
        _Posted = value
      End Set
    End Property
    Public Property Posted1LeaveTypeID() As String
      Get
        Return _Posted1LeaveTypeID
      End Get
      Set(ByVal value As String)
				 If Convert.IsDBNull(Value) Then
					 _Posted1LeaveTypeID = ""
				 Else
					 _Posted1LeaveTypeID = value
			   End If
      End Set
    End Property
    Public Property Posted2LeaveTypeID() As String
      Get
        Return _Posted2LeaveTypeID
      End Get
      Set(ByVal value As String)
				 If Convert.IsDBNull(Value) Then
					 _Posted2LeaveTypeID = ""
				 Else
					 _Posted2LeaveTypeID = value
			   End If
      End Set
    End Property
    Public Property ApplHeaderID() As String
      Get
        Return _ApplHeaderID
      End Get
      Set(ByVal value As String)
				 If Convert.IsDBNull(Value) Then
					 _ApplHeaderID = ""
				 Else
					 _ApplHeaderID = value
			   End If
      End Set
    End Property
    Public Property ApplStatusID() As String
      Get
        Return _ApplStatusID
      End Get
      Set(ByVal value As String)
				 If Convert.IsDBNull(Value) Then
					 _ApplStatusID = ""
				 Else
					 _ApplStatusID = value
			   End If
      End Set
    End Property
    Public Property FinalValue() As String
      Get
        Return _FinalValue
      End Get
      Set(ByVal value As String)
				 If Convert.IsDBNull(Value) Then
					 _FinalValue = ""
				 Else
					 _FinalValue = value
			   End If
      End Set
    End Property
    Public Property AdvanceApplication() As Boolean
      Get
        Return _AdvanceApplication
      End Get
      Set(ByVal value As Boolean)
        _AdvanceApplication = value
      End Set
    End Property
    Public Property MannuallyCorrected() As Boolean
      Get
        Return _MannuallyCorrected
      End Get
      Set(ByVal value As Boolean)
        _MannuallyCorrected = value
      End Set
    End Property
    Public Property Destination() As String
      Get
        Return _Destination
      End Get
      Set(ByVal value As String)
				 If Convert.IsDBNull(Value) Then
					 _Destination = ""
				 Else
					 _Destination = value
			   End If
      End Set
    End Property
    Public Property Purpose() As String
      Get
        Return _Purpose
      End Get
      Set(ByVal value As String)
				 If Convert.IsDBNull(Value) Then
					 _Purpose = ""
				 Else
					 _Purpose = value
			   End If
      End Set
    End Property
    Public Property ConfigID() As String
      Get
        Return _ConfigID
      End Get
      Set(ByVal value As String)
				 If Convert.IsDBNull(Value) Then
					 _ConfigID = ""
				 Else
					 _ConfigID = value
			   End If
      End Set
    End Property
    Public Property ConfigDetailID() As String
      Get
        Return _ConfigDetailID
      End Get
      Set(ByVal value As String)
				 If Convert.IsDBNull(Value) Then
					 _ConfigDetailID = ""
				 Else
					 _ConfigDetailID = value
			   End If
      End Set
    End Property
    Public Property ConfigStatus() As String
      Get
        Return _ConfigStatus
      End Get
      Set(ByVal value As String)
				 If Convert.IsDBNull(Value) Then
					 _ConfigStatus = ""
				 Else
					 _ConfigStatus = value
			   End If
      End Set
    End Property
    Public Property TSStatus() As String
      Get
        Return _TSStatus
      End Get
      Set(ByVal value As String)
				 If Convert.IsDBNull(Value) Then
					 _TSStatus = ""
				 Else
					 _TSStatus = value
			   End If
      End Set
    End Property
    Public Property TSStatusBy() As String
      Get
        Return _TSStatusBy
      End Get
      Set(ByVal value As String)
				 If Convert.IsDBNull(Value) Then
					 _TSStatusBy = ""
				 Else
					 _TSStatusBy = value
			   End If
      End Set
    End Property
    Public Property TSStatusOn() As String
      Get
        If Not _TSStatusOn = String.Empty Then
          Return Convert.ToDateTime(_TSStatusOn).ToString("dd/MM/yyyy")
        End If
        Return _TSStatusOn
      End Get
      Set(ByVal value As String)
				 If Convert.IsDBNull(Value) Then
					 _TSStatusOn = ""
				 Else
					 _TSStatusOn = value
			   End If
      End Set
    End Property
    <DataObjectMethod(DataObjectMethodType.Select)> _
    Public Shared Function GetByID(ByVal AttenID As Int32) As SIS.ATN.atnNewAttendance
      Dim Results As SIS.ATN.atnNewAttendance = Nothing
      Using Con As SqlConnection = New SqlConnection(SIS.SYS.SQLDatabase.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "spatnNewAttendanceSelectByID"
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@AttenID",SqlDbType.Int,AttenID.ToString.Length, AttenID)
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          If Reader.Read() Then
            Results = New SIS.ATN.atnNewAttendance(Reader)
          End If
          Reader.Close()
        End Using
      End Using
      Return Results
    End Function
      'Select By ID One Record Filtered Overloaded GetByID
    <DataObjectMethod(DataObjectMethodType.Select)> _
    Public Shared Function SelectList(ByVal startRowIndex As Integer, ByVal maximumRows As Integer, ByVal orderBy As String, ByVal SearchState As Boolean, ByVal SearchText As String, ByVal FinYear As String) As List(Of SIS.ATN.atnNewAttendance)
      Dim Results As List(Of SIS.ATN.atnNewAttendance) = Nothing
      Using Con As SqlConnection = New SqlConnection(SIS.SYS.SQLDatabase.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          If orderBy = String.Empty Then orderBy = "AttenDate DESC"
          Cmd.CommandType = CommandType.StoredProcedure
          If SearchState Then
            Cmd.CommandText = "spatnNewAttendanceSelectListSearch"
            SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@KeyWord", SqlDbType.NVarChar, 250, SearchText)
          Else
            Cmd.CommandText = "spatnNewAttendanceSelectListFilteres"
          End If
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@startRowIndex", SqlDbType.Int, -1, startRowIndex)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@maximumRows", SqlDbType.Int, -1, maximumRows)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@OrderBy", SqlDbType.NVarChar, 50, orderBy)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@FinYear", SqlDbType.NVarChar, 4, FinYear)
          Cmd.Parameters.Add("@RecordCount", SqlDbType.Int)
          Cmd.Parameters("@RecordCount").Direction = ParameterDirection.Output
          _RecordCount = -1
          Results = New List(Of SIS.ATN.atnNewAttendance)()
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          While (Reader.Read())
            Results.Add(New SIS.ATN.atnNewAttendance(Reader))
          End While
          Reader.Close()
          _RecordCount = Cmd.Parameters("@RecordCount").Value
        End Using
      End Using
      Return Results
    End Function
    <DataObjectMethod(DataObjectMethodType.Insert, True)> _
    Public Shared Function Insert(ByVal Record As SIS.ATN.atnNewAttendance) As Int32
      Dim _Result As Int32 = Record.AttenID
			SIS.ATN.atnNewAttendance.SetPunch9Time(Record)
			Using Con As SqlConnection = New SqlConnection(SIS.SYS.SQLDatabase.DBCommon.GetConnectionString())
				Using Cmd As SqlCommand = Con.CreateCommand()
					Cmd.CommandType = CommandType.StoredProcedure
					Cmd.CommandText = "spatnNewAttendanceInsert"
					SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@AttenDate", SqlDbType.DateTime, 21, Record.AttenDate)
					SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@CardNo", SqlDbType.NVarChar, 9, Record.CardNo)
					SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@Punch1Time", SqlDbType.Decimal, 9, IIf(Record.Punch1Time = "", Convert.DBNull, Record.Punch1Time))
					SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@Punch2Time", SqlDbType.Decimal, 9, IIf(Record.Punch2Time = "", Convert.DBNull, Record.Punch2Time))
					SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@PunchStatusID", SqlDbType.NVarChar, 3, IIf(Record.PunchStatusID = "", Convert.DBNull, Record.PunchStatusID))
					SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@Punch9Time", SqlDbType.Decimal, 9, IIf(Record.Punch9Time = "", Convert.DBNull, Record.Punch9Time))
					SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@PunchValue", SqlDbType.Decimal, 9, IIf(Record.PunchValue = "", Convert.DBNull, Record.PunchValue))
					SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@NeedsRegularization", SqlDbType.Bit, 3, Record.NeedsRegularization)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@FinYear", SqlDbType.NVarChar, 5, Record.FinYear)
					SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@Applied", SqlDbType.Bit, 3, Record.Applied)
					SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@AppliedValue", SqlDbType.Decimal, 9, IIf(Record.AppliedValue = "", Convert.DBNull, Record.AppliedValue))
					SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@Applied1LeaveTypeID", SqlDbType.NVarChar, 3, IIf(Record.Applied1LeaveTypeID = "", Convert.DBNull, Record.Applied1LeaveTypeID))
					SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@Applied2LeaveTypeID", SqlDbType.NVarChar, 3, IIf(Record.Applied2LeaveTypeID = "", Convert.DBNull, Record.Applied2LeaveTypeID))
					SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@Posted", SqlDbType.Bit, 3, Record.Posted)
					SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@Posted1LeaveTypeID", SqlDbType.NVarChar, 3, IIf(Record.Posted1LeaveTypeID = "", Convert.DBNull, Record.Posted1LeaveTypeID))
					SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@Posted2LeaveTypeID", SqlDbType.NVarChar, 3, IIf(Record.Posted2LeaveTypeID = "", Convert.DBNull, Record.Posted2LeaveTypeID))
					SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@ApplHeaderID", SqlDbType.Int, 11, IIf(Record.ApplHeaderID = "", Convert.DBNull, Record.ApplHeaderID))
					SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@ApplStatusID", SqlDbType.Int, 11, IIf(Record.ApplStatusID = "", Convert.DBNull, Record.ApplStatusID))
					SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@FinalValue", SqlDbType.Decimal, 9, IIf(Record.FinalValue = "", Convert.DBNull, Record.FinalValue))
					SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@AdvanceApplication", SqlDbType.Bit, 3, Record.AdvanceApplication)
					SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@MannuallyCorrected", SqlDbType.Bit, 3, Record.MannuallyCorrected)
					SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@Destination", SqlDbType.NVarChar, 51, IIf(Record.Destination = "", Convert.DBNull, Record.Destination))
					SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@Purpose", SqlDbType.NVarChar, 251, IIf(Record.Purpose = "", Convert.DBNull, Record.Purpose))
					SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@ConfigID", SqlDbType.Int, 11, IIf(Record.ConfigID = "", Convert.DBNull, Record.ConfigID))
					SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@ConfigDetailID", SqlDbType.Int, 11, IIf(Record.ConfigDetailID = "", Convert.DBNull, Record.ConfigDetailID))
					SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@ConfigStatus", SqlDbType.NVarChar, 3, IIf(Record.ConfigStatus = "", Convert.DBNull, Record.ConfigStatus))
					SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@TSStatus", SqlDbType.NVarChar, 3, IIf(Record.TSStatus = "", Convert.DBNull, Record.TSStatus))
					SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@TSStatusBy", SqlDbType.NVarChar, 9, IIf(Record.TSStatusBy = "", Convert.DBNull, Record.TSStatusBy))
					SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@TSStatusOn", SqlDbType.DateTime, 21, IIf(Record.TSStatusOn = "", Convert.DBNull, Record.TSStatusOn))
					SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@HoliDay", SqlDbType.Bit, 3, Record.HoliDay)
					SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@SiteAttendance", SqlDbType.Bit, 3, Record.SiteAttendance)
					SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@SiteAttendanceVerified", SqlDbType.Bit, 3, Record.SiteAttendanceVerified)
					SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@SiteAttendanceVerifiedBy", SqlDbType.NVarChar, 9, IIf(Record.SiteAttendanceVerifiedBy = "", Convert.DBNull, Record.SiteAttendanceVerifiedBy))
					SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@SiteAttendanceVerifiedOn", SqlDbType.DateTime, 21, IIf(Record.SiteAttendanceVerifiedOn = "", Convert.DBNull, Record.SiteAttendanceVerifiedOn))
					SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@FirstPunchMachine", SqlDbType.NVarChar, 100, IIf(Record.FirstPunchMachine = "", Convert.DBNull, Record.FirstPunchMachine))
					SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@SecondPunchMachine", SqlDbType.NVarChar, 100, IIf(Record.SecondPunchMachine = "", Convert.DBNull, Record.SecondPunchMachine))
					Cmd.Parameters.Add("@Return_AttenID", SqlDbType.Int, 10)
					Cmd.Parameters("@Return_AttenID").Direction = ParameterDirection.Output
					Con.Open()
					Cmd.ExecuteNonQuery()
					_Result = Cmd.Parameters("@Return_AttenID").Value
				End Using
			End Using
      Return _Result
    End Function
    <DataObjectMethod(DataObjectMethodType.Update, True)> _
    Public Shared Function Update(ByVal Record As SIS.ATN.atnNewAttendance) As Int32
      Dim _Result as Integer = 0
      Using Con As SqlConnection = New SqlConnection(SIS.SYS.SQLDatabase.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "spatnNewAttendanceUpdate"
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@Original_AttenID",SqlDbType.Int,11, Record.AttenID)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@AttenDate",SqlDbType.DateTime,21, Record.AttenDate)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@CardNo",SqlDbType.NVarChar,9, Record.CardNo)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@Punch1Time",SqlDbType.Decimal,9, Iif(Record.Punch1Time= "" ,Convert.DBNull, Record.Punch1Time))
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@Punch2Time",SqlDbType.Decimal,9, Iif(Record.Punch2Time= "" ,Convert.DBNull, Record.Punch2Time))
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@PunchStatusID",SqlDbType.NVarChar,3, Iif(Record.PunchStatusID= "" ,Convert.DBNull, Record.PunchStatusID))
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@Punch9Time",SqlDbType.Decimal,9, Iif(Record.Punch9Time= "" ,Convert.DBNull, Record.Punch9Time))
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@PunchValue",SqlDbType.Decimal,9, Iif(Record.PunchValue= "" ,Convert.DBNull, Record.PunchValue))
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@NeedsRegularization",SqlDbType.Bit,3, Record.NeedsRegularization)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@Applied",SqlDbType.Bit,3, Record.Applied)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@AppliedValue",SqlDbType.Decimal,9, Iif(Record.AppliedValue= "" ,Convert.DBNull, Record.AppliedValue))
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@Applied1LeaveTypeID",SqlDbType.NVarChar,3, Iif(Record.Applied1LeaveTypeID= "" ,Convert.DBNull, Record.Applied1LeaveTypeID))
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@Applied2LeaveTypeID",SqlDbType.NVarChar,3, Iif(Record.Applied2LeaveTypeID= "" ,Convert.DBNull, Record.Applied2LeaveTypeID))
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@Posted",SqlDbType.Bit,3, Record.Posted)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@Posted1LeaveTypeID",SqlDbType.NVarChar,3, Iif(Record.Posted1LeaveTypeID= "" ,Convert.DBNull, Record.Posted1LeaveTypeID))
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@Posted2LeaveTypeID",SqlDbType.NVarChar,3, Iif(Record.Posted2LeaveTypeID= "" ,Convert.DBNull, Record.Posted2LeaveTypeID))
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@ApplHeaderID",SqlDbType.Int,11, Iif(Record.ApplHeaderID= "" ,Convert.DBNull, Record.ApplHeaderID))
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@ApplStatusID",SqlDbType.Int,11, Iif(Record.ApplStatusID= "" ,Convert.DBNull, Record.ApplStatusID))
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@FinalValue",SqlDbType.Decimal,9, Iif(Record.FinalValue= "" ,Convert.DBNull, Record.FinalValue))
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@AdvanceApplication",SqlDbType.Bit,3, Record.AdvanceApplication)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@MannuallyCorrected",SqlDbType.Bit,3, Record.MannuallyCorrected)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@Destination",SqlDbType.NVarChar,51, Iif(Record.Destination= "" ,Convert.DBNull, Record.Destination))
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@Purpose",SqlDbType.NVarChar,251, Iif(Record.Purpose= "" ,Convert.DBNull, Record.Purpose))
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@ConfigID",SqlDbType.Int,11, Iif(Record.ConfigID= "" ,Convert.DBNull, Record.ConfigID))
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@ConfigDetailID",SqlDbType.Int,11, Iif(Record.ConfigDetailID= "" ,Convert.DBNull, Record.ConfigDetailID))
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@ConfigStatus",SqlDbType.NVarChar,3, Iif(Record.ConfigStatus= "" ,Convert.DBNull, Record.ConfigStatus))
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@TSStatus",SqlDbType.NVarChar,3, Iif(Record.TSStatus= "" ,Convert.DBNull, Record.TSStatus))
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@TSStatusBy",SqlDbType.NVarChar,9, Iif(Record.TSStatusBy= "" ,Convert.DBNull, Record.TSStatusBy))
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@TSStatusOn",SqlDbType.DateTime,21, Iif(Record.TSStatusOn= "" ,Convert.DBNull, Record.TSStatusOn))
					SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@HoliDay", SqlDbType.Bit, 3, Record.HoliDay)
					SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@SiteAttendance", SqlDbType.Bit, 3, Record.SiteAttendance)
					SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@SiteAttendanceVerified", SqlDbType.Bit, 3, Record.SiteAttendanceVerified)
					SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@SiteAttendanceVerifiedBy", SqlDbType.NVarChar, 9, IIf(Record.SiteAttendanceVerifiedBy = "", Convert.DBNull, Record.SiteAttendanceVerifiedBy))
					SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@SiteAttendanceVerifiedOn", SqlDbType.DateTime, 21, IIf(Record.SiteAttendanceVerifiedOn = "", Convert.DBNull, Record.SiteAttendanceVerifiedOn))
					SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@FirstPunchMachine", SqlDbType.NVarChar, 100, IIf(Record.FirstPunchMachine = "", Convert.DBNull, Record.FirstPunchMachine))
					SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@SecondPunchMachine", SqlDbType.NVarChar, 100, IIf(Record.SecondPunchMachine = "", Convert.DBNull, Record.SecondPunchMachine))
					Cmd.Parameters.Add("@RowCount", SqlDbType.Int)
          Cmd.Parameters("@RowCount").Direction = ParameterDirection.Output
					Con.Open()
					Try
						Cmd.ExecuteNonQuery()
					Catch ex As Exception

					End Try
					_Result = Cmd.Parameters("@RowCount").Value
        End Using
      End Using
      Return _Result
    End Function
    <DataObjectMethod(DataObjectMethodType.Delete, True)> _
    Public Shared Function Delete(ByVal Record As SIS.ATN.atnNewAttendance) As Int32
      Dim _Result as Integer = 0
      Using Con As SqlConnection = New SqlConnection(SIS.SYS.SQLDatabase.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "spatnNewAttendanceDelete"
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@Original_AttenID",SqlDbType.Int,Record.AttenID.ToString.Length, Record.AttenID)
          Cmd.Parameters.Add("@RowCount", SqlDbType.Int)
          Cmd.Parameters("@RowCount").Direction = ParameterDirection.Output
          Con.Open()
          Cmd.ExecuteNonQuery()
          _Result = Cmd.Parameters("@RowCount").Value
        End Using
      End Using
      Return _Result
    End Function
    Public Shared Function SelectCount(ByVal SearchState As Boolean, ByVal SearchText As String) As Integer
      Return _RecordCount
    End Function
		Public Sub New(ByVal Reader As SqlDataReader)
			Try
				For Each pi As System.Reflection.PropertyInfo In Me.GetType.GetProperties
					If pi.MemberType = Reflection.MemberTypes.Property Then
						Try
							If Convert.IsDBNull(Reader(pi.Name)) Then
								CallByName(Me, pi.Name, CallType.Let, String.Empty)
							Else
								CallByName(Me, pi.Name, CallType.Let, Reader(pi.Name))
							End If
						Catch ex As Exception
						End Try
					End If
				Next
			Catch ex As Exception
			End Try

    End Sub
    Public Sub New()
    End Sub
  End Class
End Namespace

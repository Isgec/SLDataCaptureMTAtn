Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Namespace SIS.ATN
  <DataObject()> _
  Partial Public Class atnRawPunch
    Private Shared _RecordCount As Integer
    Private _RecordID As Int32
    Private _PunchDate As String
    Private _CardNo As String
    Private _Punch1Time As Decimal
    Private _Punch2Time As Decimal
    Private _Processed As Boolean
    Private _FinYear As String
    Private _CardNoEmployeeName As String
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
    Public Property RecordID() As Int32
      Get
        Return _RecordID
      End Get
      Set(ByVal value As Int32)
        _RecordID = value
      End Set
    End Property
    Public Property PunchDate() As String
      Get
        If Not _PunchDate = String.Empty Then
          Return Convert.ToDateTime(_PunchDate).ToString("dd/MM/yyyy")
        End If
        Return _PunchDate
      End Get
      Set(ByVal value As String)
			   _PunchDate = value
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
    Public Property Punch1Time() As Decimal
      Get
        Return _Punch1Time
      End Get
      Set(ByVal value As Decimal)
        _Punch1Time = value
      End Set
    End Property
    Public Property Punch2Time() As Decimal
      Get
        Return _Punch2Time
      End Get
      Set(ByVal value As Decimal)
        _Punch2Time = value
      End Set
    End Property
    Public Property Processed() As Boolean
      Get
        Return _Processed
      End Get
      Set(ByVal value As Boolean)
        _Processed = value
      End Set
    End Property
    Public Property FinYear() As String
      Get
        Return _FinYear
      End Get
      Set(ByVal value As String)
        _FinYear = value
      End Set
    End Property
    Public Property CardNoEmployeeName() As String
      Get
        Return _CardNoEmployeeName
      End Get
      Set(ByVal value As String)
        _CardNoEmployeeName = value
      End Set
    End Property
    <DataObjectMethod(DataObjectMethodType.Select)> _
    Public Shared Function GetByID(ByVal RecordID As Int32) As SIS.ATN.atnRawPunch
      Dim Results As SIS.ATN.atnRawPunch = Nothing
      Using Con As SqlConnection = New SqlConnection(SIS.SYS.SQLDatabase.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "spatnRawPunchSelectByID"
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@RecordID",SqlDbType.Int,RecordID.ToString.Length, RecordID)
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          If Reader.Read() Then
            Results = New SIS.ATN.atnRawPunch(Reader)
          End If
          Reader.Close()
        End Using
      End Using
      Return Results
    End Function
      'Select By ID One Record Filtered Overloaded GetByID
    <DataObjectMethod(DataObjectMethodType.Select)> _
    Public Shared Function GetByID(ByVal RecordID As Int32, ByVal CardNo As String) As SIS.ATN.atnRawPunch
      Return GetByID(RecordID)
    End Function
    <DataObjectMethod(DataObjectMethodType.Insert, True)> _
    Public Shared Function Insert(ByVal Record As SIS.ATN.atnRawPunch) As Int32
      Dim _Result As Int32 = Record.RecordID
      Using Con As SqlConnection = New SqlConnection(SIS.SYS.SQLDatabase.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "spatnRawPunchInsert"
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@PunchDate", SqlDbType.DateTime, 21, Record.PunchDate)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@CardNo", SqlDbType.NVarChar, 9, Record.CardNo)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@Punch1Time", SqlDbType.Decimal, 9, Record.Punch1Time)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@Punch2Time", SqlDbType.Decimal, 9, Record.Punch2Time)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@Processed", SqlDbType.Bit, 3, Record.Processed)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@FinYear", SqlDbType.NVarChar, 5, Record.FinYear)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@FirstPunchMachine", SqlDbType.NVarChar, 100, IIf(Record.FirstPunchMachine = "", Convert.DBNull, Record.FirstPunchMachine))
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@SecondPunchMachine", SqlDbType.NVarChar, 100, IIf(Record.SecondPunchMachine = "", Convert.DBNull, Record.SecondPunchMachine))
          Cmd.Parameters.Add("@Return_RecordID", SqlDbType.Int, 10)
          Cmd.Parameters("@Return_RecordID").Direction = ParameterDirection.Output
          Con.Open()
          Cmd.ExecuteNonQuery()
          _Result = Cmd.Parameters("@Return_RecordID").Value
        End Using
      End Using
      Return _Result
    End Function
    <DataObjectMethod(DataObjectMethodType.Update, True)> _
    Public Shared Function Update(ByVal Record As SIS.ATN.atnRawPunch) As Int32
      Dim _Result as Integer = 0
      Using Con As SqlConnection = New SqlConnection(SIS.SYS.SQLDatabase.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "spatnRawPunchUpdate"
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@Original_RecordID",SqlDbType.Int,11, Record.RecordID)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@PunchDate",SqlDbType.DateTime,21, Record.PunchDate)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@CardNo",SqlDbType.NVarChar,9, Record.CardNo)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@Punch1Time",SqlDbType.Decimal,9, Record.Punch1Time)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@Punch2Time",SqlDbType.Decimal,9, Record.Punch2Time)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@Processed",SqlDbType.Bit,3, Record.Processed)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@FinYear", SqlDbType.NVarChar, 5, Record.FinYear)
					SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@FirstPunchMachine", SqlDbType.NVarChar, 100, IIf(Record.FirstPunchMachine = "", Convert.DBNull, Record.FirstPunchMachine))
					SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@SecondPunchMachine", SqlDbType.NVarChar, 100, IIf(Record.SecondPunchMachine = "", Convert.DBNull, Record.SecondPunchMachine))
					Cmd.Parameters.Add("@RowCount", SqlDbType.Int)
          Cmd.Parameters("@RowCount").Direction = ParameterDirection.Output
          Con.Open()
          Cmd.ExecuteNonQuery()
          _Result = Cmd.Parameters("@RowCount").Value
        End Using
      End Using
      Return _Result
    End Function
    <DataObjectMethod(DataObjectMethodType.Delete, True)> _
    Public Shared Function Delete(ByVal Record As SIS.ATN.atnRawPunch) As Int32
      Dim _Result as Integer = 0
      Using Con As SqlConnection = New SqlConnection(SIS.SYS.SQLDatabase.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "spatnRawPunchDelete"
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@Original_RecordID",SqlDbType.Int,Record.RecordID.ToString.Length, Record.RecordID)
          Cmd.Parameters.Add("@RowCount", SqlDbType.Int)
          Cmd.Parameters("@RowCount").Direction = ParameterDirection.Output
          Con.Open()
          Cmd.ExecuteNonQuery()
          _Result = Cmd.Parameters("@RowCount").Value
        End Using
      End Using
      Return _Result
    End Function
    Public Shared Function SelectCount(ByVal SearchState As Boolean, ByVal SearchText As String, ByVal CardNo As String) As Integer
      Return _RecordCount
    End Function
    Public Sub New(ByVal Reader As SqlDataReader)
      On Error Resume Next
      _RecordID = Ctype(Reader("RecordID"),Int32)
      _PunchDate = Ctype(Reader("PunchDate"),DateTime)
      _CardNo = Ctype(Reader("CardNo"),String)
      _Punch1Time = Ctype(Reader("Punch1Time"),Decimal)
      _Punch2Time = Ctype(Reader("Punch2Time"),Decimal)
      _Processed = Ctype(Reader("Processed"),Boolean)
			_FinYear = CType(Reader("FinYear"), String)
			If Convert.IsDBNull(Reader("FirstPunchMachine")) Then
				_FirstPunchMachine = String.Empty
			Else
				_FirstPunchMachine = CType(Reader("FirstPunchMachine"), String)
			End If
			If Convert.IsDBNull(Reader("SecondPunchMachine")) Then
				_SecondPunchMachine = String.Empty
			Else
				_SecondPunchMachine = CType(Reader("SecondPunchMachine"), String)
			End If
			_CardNoEmployeeName = Reader("HRM_Employees1_EmployeeName") & " [" & CType(Reader("CardNo"), String) & "]"
    End Sub
    Public Sub New()
    End Sub
  End Class
End Namespace

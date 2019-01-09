Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Namespace SIS.ATN
  <DataObject()> _
  Partial Public Class atnPunchConfig
    Private Shared _RecordCount As Integer
    Private _RecordID As Int32
    Private _Description As String
    Private _STD1Time As Decimal
    Private _Range1Start As Decimal
    Private _Range1End As Decimal
    Private _MeanTime As Decimal
    Private _STD2Time As Decimal
    Private _Range2Start As Decimal
    Private _Range2End As Decimal
    Private _EnableMinHrs As Boolean
    Private _MinHrsFullPresent As Decimal
    Private _MinHrsHalfPresent As Decimal
    Private _Active As Boolean
    Private _FinYear As String
    Private _DataFileLocation As String
    Public Property RecordID() As Int32
      Get
        Return _RecordID
      End Get
      Set(ByVal value As Int32)
        _RecordID = value
      End Set
    End Property
    Public Property Description() As String
      Get
        Return _Description
      End Get
      Set(ByVal value As String)
        _Description = value
      End Set
    End Property
    Public Property STD1Time() As Decimal
      Get
        Return _STD1Time
      End Get
      Set(ByVal value As Decimal)
        _STD1Time = value
      End Set
    End Property
    Public Property Range1Start() As Decimal
      Get
        Return _Range1Start
      End Get
      Set(ByVal value As Decimal)
        _Range1Start = value
      End Set
    End Property
    Public Property Range1End() As Decimal
      Get
        Return _Range1End
      End Get
      Set(ByVal value As Decimal)
        _Range1End = value
      End Set
    End Property
    Public Property MeanTime() As Decimal
      Get
        Return _MeanTime
      End Get
      Set(ByVal value As Decimal)
        _MeanTime = value
      End Set
    End Property
    Public Property STD2Time() As Decimal
      Get
        Return _STD2Time
      End Get
      Set(ByVal value As Decimal)
        _STD2Time = value
      End Set
    End Property
    Public Property Range2Start() As Decimal
      Get
        Return _Range2Start
      End Get
      Set(ByVal value As Decimal)
        _Range2Start = value
      End Set
    End Property
    Public Property Range2End() As Decimal
      Get
        Return _Range2End
      End Get
      Set(ByVal value As Decimal)
        _Range2End = value
      End Set
    End Property
    Public Property EnableMinHrs() As Boolean
      Get
        Return _EnableMinHrs
      End Get
      Set(ByVal value As Boolean)
        _EnableMinHrs = value
      End Set
    End Property
    Public Property MinHrsFullPresent() As Decimal
      Get
        Return _MinHrsFullPresent
      End Get
      Set(ByVal value As Decimal)
        _MinHrsFullPresent = value
      End Set
    End Property
    Public Property MinHrsHalfPresent() As Decimal
      Get
        Return _MinHrsHalfPresent
      End Get
      Set(ByVal value As Decimal)
        _MinHrsHalfPresent = value
      End Set
    End Property
    Public Property Active() As Boolean
      Get
        Return _Active
      End Get
      Set(ByVal value As Boolean)
        _Active = value
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
    Public Property DataFileLocation() As String
      Get
        Return _DataFileLocation
      End Get
      Set(ByVal value As String)
        _DataFileLocation = value
      End Set
    End Property
    <DataObjectMethod(DataObjectMethodType.Select)> _
    Public Shared Function SelectList(ByVal orderBy As String, ByVal FinYear As String) As List(Of SIS.ATN.atnPunchConfig)
      Dim Results As List(Of SIS.ATN.atnPunchConfig) = Nothing
      Using Con As SqlConnection = New SqlConnection(SIS.SYS.SQLDatabase.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "spatnPunchConfigSelectList"
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@OrderBy", SqlDbType.NVarChar, 50, orderBy)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@FinYear", SqlDbType.NVarChar, 4, FinYear)
          Cmd.Parameters.Add("@RecordCount", SqlDbType.Int)
          Cmd.Parameters("@RecordCount").Direction = ParameterDirection.Output
          _RecordCount = -1
          Results = New List(Of SIS.ATN.atnPunchConfig)()
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          While (Reader.Read())
            Results.Add(New SIS.ATN.atnPunchConfig(Reader))
          End While
          Reader.Close()
          _RecordCount = Cmd.Parameters("@RecordCount").Value
        End Using
      End Using
      Return Results
    End Function
    <DataObjectMethod(DataObjectMethodType.Select)> _
    Public Shared Function GetByID(ByVal RecordID As Int32) As SIS.ATN.atnPunchConfig
      Dim Results As SIS.ATN.atnPunchConfig = Nothing
      Using Con As SqlConnection = New SqlConnection(SIS.SYS.SQLDatabase.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "spatnPunchConfigSelectByID"
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@RecordID",SqlDbType.Int,RecordID.ToString.Length, RecordID)
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          If Reader.Read() Then
            Results = New SIS.ATN.atnPunchConfig(Reader)
          End If
          Reader.Close()
        End Using
      End Using
      Return Results
    End Function
      'Select By ID One Record Filtered Overloaded GetByID
    <DataObjectMethod(DataObjectMethodType.Select)> _
    Public Shared Function SelectList(ByVal startRowIndex As Integer, ByVal maximumRows As Integer, ByVal orderBy As String, ByVal SearchState As Boolean, ByVal SearchText As String, ByVal FinYear As String) As List(Of SIS.ATN.atnPunchConfig)
      Dim Results As List(Of SIS.ATN.atnPunchConfig) = Nothing
      Using Con As SqlConnection = New SqlConnection(SIS.SYS.SQLDatabase.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          If SearchState Then
            Cmd.CommandText = "spatnPunchConfigSelectListSearch"
            SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@KeyWord", SqlDbType.NVarChar, 250, SearchText)
          Else
            Cmd.CommandText = "spatnPunchConfigSelectListFilteres"
          End If
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@startRowIndex", SqlDbType.Int, -1, startRowIndex)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@maximumRows", SqlDbType.Int, -1, maximumRows)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@OrderBy", SqlDbType.NVarChar, 50, orderBy)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@FinYear", SqlDbType.NVarChar, 4, FinYear)
          Cmd.Parameters.Add("@RecordCount", SqlDbType.Int)
          Cmd.Parameters("@RecordCount").Direction = ParameterDirection.Output
          _RecordCount = -1
          Results = New List(Of SIS.ATN.atnPunchConfig)()
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          While (Reader.Read())
            Results.Add(New SIS.ATN.atnPunchConfig(Reader))
          End While
          Reader.Close()
          _RecordCount = Cmd.Parameters("@RecordCount").Value
        End Using
      End Using
      Return Results
    End Function
    Public Shared Function SelectCount(ByVal SearchState As Boolean, ByVal SearchText As String) As Integer
      Return _RecordCount
    End Function
'		Autocomplete Method
    Public Sub New(ByVal Reader As SqlDataReader)
      On Error Resume Next
      _RecordID = CType(Reader("RecordID"), Int32)
      _Description = CType(Reader("Description"), String)
      _STD1Time = CType(Reader("STD1Time"), Decimal)
      _Range1Start = CType(Reader("Range1Start"), Decimal)
      _Range1End = CType(Reader("Range1End"), Decimal)
      _MeanTime = CType(Reader("MeanTime"), Decimal)
      _STD2Time = CType(Reader("STD2Time"), Decimal)
      _Range2Start = CType(Reader("Range2Start"), Decimal)
      _Range2End = CType(Reader("Range2End"), Decimal)
      _EnableMinHrs = CType(Reader("EnableMinHrs"), Boolean)
      _MinHrsFullPresent = CType(Reader("MinHrsFullPresent"), Decimal)
      _MinHrsHalfPresent = CType(Reader("MinHrsHalfPresent"), Decimal)
      _Active = CType(Reader("Active"), Boolean)
      _FinYear = CType(Reader("FinYear"), String)
      _DataFileLocation = CType(Reader("DataFileLocation"), String)
    End Sub
    Public Sub New(ByVal AliasName As String, ByVal Reader As SqlDataReader)
      On Error Resume Next
      _RecordID = Ctype(Reader(AliasName & "_RecordID"),Int32)
      _Description = Ctype(Reader(AliasName & "_Description"),String)
      _STD1Time = Ctype(Reader(AliasName & "_STD1Time"),Decimal)
      _Range1Start = Ctype(Reader(AliasName & "_Range1Start"),Decimal)
      _Range1End = Ctype(Reader(AliasName & "_Range1End"),Decimal)
      _MeanTime = Ctype(Reader(AliasName & "_MeanTime"),Decimal)
      _STD2Time = Ctype(Reader(AliasName & "_STD2Time"),Decimal)
      _Range2Start = Ctype(Reader(AliasName & "_Range2Start"),Decimal)
      _Range2End = Ctype(Reader(AliasName & "_Range2End"),Decimal)
      _EnableMinHrs = Ctype(Reader(AliasName & "_EnableMinHrs"),Boolean)
      _MinHrsFullPresent = Ctype(Reader(AliasName & "_MinHrsFullPresent"),Decimal)
      _MinHrsHalfPresent = Ctype(Reader(AliasName & "_MinHrsHalfPresent"),Decimal)
      _Active = Ctype(Reader(AliasName & "_Active"),Boolean)
      _FinYear = Ctype(Reader(AliasName & "_FinYear"),String)
      _DataFileLocation = Ctype(Reader(AliasName & "_DataFileLocation"),String)
    End Sub
    Public Sub New()
    End Sub
  End Class
End Namespace

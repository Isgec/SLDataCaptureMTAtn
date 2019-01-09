Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Namespace SIS.ATN
  <DataObject()> _
  Partial Public Class atnHolidays
    Private Shared _RecordCount As Integer
    Private _RecordID As Int32
    Private _Holiday As String
    Private _OfficeID As String
    Private _Description As String
    Private _PunchStatusID As String
    Private _FinYear As String
    Public Property RecordID() As Int32
      Get
        Return _RecordID
      End Get
      Set(ByVal value As Int32)
        _RecordID = value
      End Set
    End Property
    Public Property Holiday() As String
      Get
        If Not _Holiday = String.Empty Then
          Return Convert.ToDateTime(_Holiday).ToString("dd/MM/yyyy")
        End If
        Return _Holiday
      End Get
      Set(ByVal value As String)
			   _Holiday = value
      End Set
    End Property
    Public Property OfficeID() As String
      Get
        Return _OfficeID
      End Get
      Set(ByVal value As String)
				 If Convert.IsDBNull(Value) Then
					 _OfficeID = ""
				 Else
					 _OfficeID = value
			   End If
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
    Public Property PunchStatusID() As String
      Get
        Return _PunchStatusID
      End Get
      Set(ByVal value As String)
        _PunchStatusID = value
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
    <DataObjectMethod(DataObjectMethodType.Select)> _
    Public Shared Function GetByID(ByVal RecordID As Int32) As SIS.ATN.atnHolidays
      Dim Results As SIS.ATN.atnHolidays = Nothing
      Using Con As SqlConnection = New SqlConnection(SIS.SYS.SQLDatabase.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "spatnHolidaysSelectByID"
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@RecordID", SqlDbType.Int, RecordID.ToString.Length, RecordID)
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          If Reader.Read() Then
            Results = New SIS.ATN.atnHolidays(Reader)
          End If
          Reader.Close()
        End Using
      End Using
      Return Results
    End Function
      'Select By ID One Record Filtered Overloaded GetByID
    <DataObjectMethod(DataObjectMethodType.Select)> _
    Public Shared Function GetByID(ByVal RecordID As Int32, ByVal OfficeID As Int32) As SIS.ATN.atnHolidays
      Return GetByID(RecordID)
    End Function
    <DataObjectMethod(DataObjectMethodType.Select)> _
    Public Shared Function SelectList(ByVal startRowIndex As Integer, ByVal maximumRows As Integer, ByVal orderBy As String, ByVal SearchState As Boolean, ByVal SearchText As String, ByVal OfficeID As Int32, ByVal FinYear As String) As List(Of SIS.ATN.atnHolidays)
      Dim Results As List(Of SIS.ATN.atnHolidays) = Nothing
      Using Con As SqlConnection = New SqlConnection(SIS.SYS.SQLDatabase.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          If SearchState Then
            Cmd.CommandText = "spatnHolidaysSelectListSearch"
            SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@KeyWord", SqlDbType.NVarChar, 250, SearchText)
          Else
            Cmd.CommandText = "spatnHolidaysSelectListFilteres"
            SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@Filter_OfficeID", SqlDbType.Int, 10, IIf(OfficeID = Nothing, 0, OfficeID))
          End If
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@startRowIndex", SqlDbType.Int, -1, startRowIndex)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@maximumRows", SqlDbType.Int, -1, maximumRows)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@OrderBy", SqlDbType.NVarChar, 50, orderBy)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@FinYear", SqlDbType.NVarChar, 4, FinYear)
          Cmd.Parameters.Add("@RecordCount", SqlDbType.Int)
          Cmd.Parameters("@RecordCount").Direction = ParameterDirection.Output
          _RecordCount = -1
          Results = New List(Of SIS.ATN.atnHolidays)()
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          While (Reader.Read())
            Results.Add(New SIS.ATN.atnHolidays(Reader))
          End While
          Reader.Close()
          _RecordCount = Cmd.Parameters("@RecordCount").Value
        End Using
      End Using
      Return Results
    End Function
    Public Shared Function SelectCount(ByVal SearchState As Boolean, ByVal SearchText As String, ByVal OfficeID As Int32) As Integer
      Return _RecordCount
    End Function
    Public Sub New(ByVal Reader As SqlDataReader)
      On Error Resume Next
      _RecordID = Ctype(Reader("RecordID"),Int32)
      _Holiday = Ctype(Reader("Holiday"),DateTime)
      If Convert.IsDBNull(Reader("OfficeID")) Then
        _OfficeID = String.Empty
      Else
        _OfficeID = Ctype(Reader("OfficeID"), String)
      End If
      _Description = Ctype(Reader("Description"),String)
      _PunchStatusID = Ctype(Reader("PunchStatusID"),String)
      _FinYear = Ctype(Reader("FinYear"),String)
    End Sub
    Public Sub New()
    End Sub
  End Class
End Namespace

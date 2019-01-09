Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Namespace SIS.ATN
  <DataObject()> _
  Partial Public Class atnEmployees
    Private Shared _RecordCount As Integer
    Private _CardNo As String
    Private _EmployeeName As String
    Private _VerificationRequired As Boolean
    Private _VerifierID As String
    Private _ApprovalRequired As Boolean
    Private _ApproverID As String
    Private _C_DateOfJoining As String
    Private _C_DateOfReleaving As String
    Private _C_CompanyID As String
    Private _C_DivisionID As String
    Private _C_OfficeID As String
    Private _C_DepartmentID As String
    Private _C_DesignationID As String
    Private _ActiveState As Boolean
		Private _C_ConfirmedOn As String
		Private _Contractual As Boolean = False
    Private _VerifierIDHRM_Employees As SIS.ATN.atnEmployees
    Private _VerifierIDEmployeeName As String
    Private _ApproverIDHRM_Employees As SIS.ATN.atnEmployees
    Private _ApproverIDEmployeeName As String
    Private _C_ProjectSiteID As String = ""
		Public Property Contractual() As Boolean
			Get
				Return _Contractual
			End Get
			Set(ByVal value As Boolean)
				_Contractual = value
			End Set
		End Property
		Public Property C_ProjectSiteID() As String
			Get
				Return _C_ProjectSiteID
			End Get
			Set(ByVal value As String)
				If Convert.IsDBNull(value) Then
					_C_ProjectSiteID = ""
				Else
					_C_ProjectSiteID = value
				End If
			End Set
		End Property
		Public ReadOnly Property DisplayField() As String
			Get
				Return _EmployeeName
			End Get
		End Property
    Public Property CardNo() As String
      Get
        Return _CardNo
      End Get
      Set(ByVal value As String)
        _CardNo = value
      End Set
    End Property
    Public Property EmployeeName() As String
      Get
        Return _EmployeeName
      End Get
      Set(ByVal value As String)
        _EmployeeName = value
      End Set
    End Property
    Public Property VerificationRequired() As Boolean
      Get
        Return _VerificationRequired
      End Get
      Set(ByVal value As Boolean)
        _VerificationRequired = value
      End Set
    End Property
    Public Property VerifierID() As String
      Get
        Return _VerifierID
      End Get
      Set(ByVal value As String)
				 If Convert.IsDBNull(Value) Then
					 _VerifierID = ""
				 Else
					 _VerifierID = value
			   End If
      End Set
    End Property
    Public Property ApprovalRequired() As Boolean
      Get
        Return _ApprovalRequired
      End Get
      Set(ByVal value As Boolean)
        _ApprovalRequired = value
      End Set
    End Property
    Public Property ApproverID() As String
      Get
        Return _ApproverID
      End Get
      Set(ByVal value As String)
				 If Convert.IsDBNull(Value) Then
					 _ApproverID = ""
				 Else
					 _ApproverID = value
			   End If
      End Set
    End Property
    Public Property C_DateOfJoining() As String
      Get
        If Not _C_DateOfJoining = String.Empty Then
          Return Convert.ToDateTime(_C_DateOfJoining).ToString("dd/MM/yyyy")
        End If
        Return _C_DateOfJoining
      End Get
      Set(ByVal value As String)
				 If Convert.IsDBNull(Value) Then
					 _C_DateOfJoining = ""
				 Else
					 _C_DateOfJoining = value
			   End If
      End Set
    End Property
    Public Property C_DateOfReleaving() As String
      Get
        If Not _C_DateOfReleaving = String.Empty Then
          Return Convert.ToDateTime(_C_DateOfReleaving).ToString("dd/MM/yyyy")
        End If
        Return _C_DateOfReleaving
      End Get
      Set(ByVal value As String)
				 If Convert.IsDBNull(Value) Then
					 _C_DateOfReleaving = ""
				 Else
					 _C_DateOfReleaving = value
			   End If
      End Set
    End Property
    Public Property C_CompanyID() As String
      Get
        Return _C_CompanyID
      End Get
      Set(ByVal value As String)
				 If Convert.IsDBNull(Value) Then
					 _C_CompanyID = ""
				 Else
					 _C_CompanyID = value
			   End If
      End Set
    End Property
    Public Property C_DivisionID() As String
      Get
        Return _C_DivisionID
      End Get
      Set(ByVal value As String)
				 If Convert.IsDBNull(Value) Then
					 _C_DivisionID = ""
				 Else
					 _C_DivisionID = value
			   End If
      End Set
    End Property
    Public Property C_OfficeID() As String
      Get
        Return _C_OfficeID
      End Get
      Set(ByVal value As String)
				 If Convert.IsDBNull(Value) Then
					 _C_OfficeID = ""
				 Else
					 _C_OfficeID = value
			   End If
      End Set
    End Property
    Public Property C_DepartmentID() As String
      Get
        Return _C_DepartmentID
      End Get
      Set(ByVal value As String)
				 If Convert.IsDBNull(Value) Then
					 _C_DepartmentID = ""
				 Else
					 _C_DepartmentID = value
			   End If
      End Set
    End Property
    Public Property C_DesignationID() As String
      Get
        Return _C_DesignationID
      End Get
      Set(ByVal value As String)
				 If Convert.IsDBNull(Value) Then
					 _C_DesignationID = ""
				 Else
					 _C_DesignationID = value
			   End If
      End Set
    End Property
    Public Property ActiveState() As Boolean
      Get
        Return _ActiveState
      End Get
      Set(ByVal value As Boolean)
        _ActiveState = value
      End Set
    End Property
    Public Property C_ConfirmedOn() As String
      Get
        If Not _C_ConfirmedOn = String.Empty Then
          Return Convert.ToDateTime(_C_ConfirmedOn).ToString("dd/MM/yyyy")
        End If
        Return _C_ConfirmedOn
      End Get
      Set(ByVal value As String)
				 If Convert.IsDBNull(Value) Then
					 _C_ConfirmedOn = ""
				 Else
					 _C_ConfirmedOn = value
			   End If
      End Set
    End Property
    Public ReadOnly Property VerifierIDHRM_Employees() As SIS.ATN.atnEmployees
      Get
        If _VerifierIDHRM_Employees Is Nothing Then
          _VerifierIDHRM_Employees = SIS.ATN.atnEmployees.GetByID(_VerifierID)
        End If
        Return _VerifierIDHRM_Employees
      End Get
    End Property
    Public Property VerifierIDEmployeeName() As String
      Get
        Return _VerifierIDEmployeeName
      End Get
      Set(ByVal value As String)
        _VerifierIDEmployeeName = value
      End Set
    End Property
    Public ReadOnly Property ApproverIDHRM_Employees() As SIS.ATN.atnEmployees
      Get
        If _ApproverIDHRM_Employees Is Nothing Then
          _ApproverIDHRM_Employees = SIS.ATN.atnEmployees.GetByID(_ApproverID)
        End If
        Return _ApproverIDHRM_Employees
      End Get
    End Property
    Public Property ApproverIDEmployeeName() As String
      Get
        Return _ApproverIDEmployeeName
      End Get
      Set(ByVal value As String)
        _ApproverIDEmployeeName = value
      End Set
    End Property
    <DataObjectMethod(DataObjectMethodType.Select)> _
    Public Shared Function SelectList(ByVal orderBy As String) As List(Of SIS.ATN.atnEmployees)
      Dim Results As List(Of SIS.ATN.atnEmployees) = Nothing
      Using Con As SqlConnection = New SqlConnection(SIS.SYS.SQLDatabase.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "spatnEmployeesSelectList"
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@OrderBy", SqlDbType.NVarChar, 50, orderBy)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@ActiveState", SqlDbType.Bit, 2, 1)
          Cmd.Parameters.Add("@RecordCount", SqlDbType.Int)
          Cmd.Parameters("@RecordCount").Direction = ParameterDirection.Output
          _RecordCount = -1
          Results = New List(Of SIS.ATN.atnEmployees)()
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          While (Reader.Read())
            Results.Add(New SIS.ATN.atnEmployees(Reader))
          End While
          Reader.Close()
          _RecordCount = Cmd.Parameters("@RecordCount").Value
        End Using
      End Using
      Return Results
    End Function
		Public Shared Function GetCardNoList(ByVal orderBy As String) As List(Of String)
			Dim Results As List(Of String) = Nothing
			Using Con As SqlConnection = New SqlConnection(SIS.SYS.SQLDatabase.DBCommon.GetConnectionString())
				Using Cmd As SqlCommand = Con.CreateCommand()
					Cmd.CommandType = CommandType.StoredProcedure
					Cmd.CommandText = "spatn_LG_EmployeesCardNoSelectList"
					SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@OrderBy", SqlDbType.NVarChar, 50, orderBy)
					SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@ActiveState", SqlDbType.Bit, 2, 1)
					Cmd.Parameters.Add("@RecordCount", SqlDbType.Int)
					Cmd.Parameters("@RecordCount").Direction = ParameterDirection.Output
					_RecordCount = -1
					Results = New List(Of String)()
					Con.Open()
					Dim Reader As SqlDataReader = Cmd.ExecuteReader()
					While (Reader.Read())
						Results.Add(Reader("CardNo"))
					End While
					Reader.Close()
					_RecordCount = Cmd.Parameters("@RecordCount").Value
				End Using
			End Using
			Return Results
		End Function
		<DataObjectMethod(DataObjectMethodType.Select)> _
		Public Shared Function GetByID(ByVal CardNo As String) As SIS.ATN.atnEmployees
			Dim Results As SIS.ATN.atnEmployees = Nothing
			Using Con As SqlConnection = New SqlConnection(SIS.SYS.SQLDatabase.DBCommon.GetConnectionString())
				Using Cmd As SqlCommand = Con.CreateCommand()
					Cmd.CommandType = CommandType.StoredProcedure
					Cmd.CommandText = "spatnEmployeesSelectByID"
					SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@CardNo", SqlDbType.NVarChar, CardNo.ToString.Length, CardNo)
					Con.Open()
					Dim Reader As SqlDataReader = Cmd.ExecuteReader()
					If Reader.Read() Then
						Results = New SIS.ATN.atnEmployees(Reader)
					End If
					Reader.Close()
				End Using
			End Using
			Return Results
		End Function
      'Select By ID One Record Filtered Overloaded GetByID
    <DataObjectMethod(DataObjectMethodType.Select)> _
    Public Shared Function SelectList(ByVal startRowIndex As Integer, ByVal maximumRows As Integer, ByVal orderBy As String, ByVal SearchState As Boolean, ByVal SearchText As String) As List(Of SIS.ATN.atnEmployees)
      Dim Results As List(Of SIS.ATN.atnEmployees) = Nothing
      Using Con As SqlConnection = New SqlConnection(SIS.SYS.SQLDatabase.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
					If SearchState Then
						Cmd.CommandText = "spatnEmployeesSelectListSearch"
						SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@KeyWord", SqlDbType.NVarChar, 250, SearchText)
					Else
						Cmd.CommandText = "spatnEmployeesSelectListFilteres"
					End If
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@startRowIndex", SqlDbType.Int, -1, startRowIndex)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@maximumRows", SqlDbType.Int, -1, maximumRows)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@OrderBy", SqlDbType.NVarChar, 50, orderBy)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@ActiveState",SqlDbType.Bit,2, 1)
          Cmd.Parameters.Add("@RecordCount", SqlDbType.Int)
          Cmd.Parameters("@RecordCount").Direction = ParameterDirection.Output
          _RecordCount = -1
          Results = New List(Of SIS.ATN.atnEmployees)()
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          While (Reader.Read())
            Results.Add(New SIS.ATN.atnEmployees(Reader))
          End While
          Reader.Close()
          _RecordCount = Cmd.Parameters("@RecordCount").Value
        End Using
      End Using
      Return Results
    End Function
    <DataObjectMethod(DataObjectMethodType.Update, True)> _
    Public Shared Function Update(ByVal Record As SIS.ATN.atnEmployees) As Int32
      Dim _Result as Integer = 0
      Using Con As SqlConnection = New SqlConnection(SIS.SYS.SQLDatabase.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "spatnEmployeesUpdate"
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@Original_CardNo",SqlDbType.NVarChar,9, Record.CardNo)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@VerificationRequired",SqlDbType.Bit,3, Record.VerificationRequired)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@VerifierID",SqlDbType.NVarChar,9, Iif(Record.VerifierID= "" ,Convert.DBNull, Record.VerifierID))
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@ApprovalRequired",SqlDbType.Bit,3, Record.ApprovalRequired)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@ApproverID",SqlDbType.NVarChar,9, Iif(Record.ApproverID= "" ,Convert.DBNull, Record.ApproverID))
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
'		Autocomplete Method
    Public Sub New(ByVal Reader As SqlDataReader)
      On Error Resume Next
      _CardNo = CType(Reader("CardNo"), String)
      _EmployeeName = CType(Reader("EmployeeName"), String)
      _VerificationRequired = CType(Reader("VerificationRequired"), Boolean)
      If Convert.IsDBNull(Reader("VerifierID")) Then
        _VerifierID = String.Empty
      Else
        _VerifierID = CType(Reader("VerifierID"), String)
      End If
      _ApprovalRequired = CType(Reader("ApprovalRequired"), Boolean)
      If Convert.IsDBNull(Reader("ApproverID")) Then
        _ApproverID = String.Empty
      Else
        _ApproverID = CType(Reader("ApproverID"), String)
      End If
      If Convert.IsDBNull(Reader("C_DateOfJoining")) Then
        _C_DateOfJoining = String.Empty
      Else
        _C_DateOfJoining = CType(Reader("C_DateOfJoining"), String)
      End If
      If Convert.IsDBNull(Reader("C_DateOfReleaving")) Then
        _C_DateOfReleaving = String.Empty
      Else
        _C_DateOfReleaving = CType(Reader("C_DateOfReleaving"), String)
      End If
      If Convert.IsDBNull(Reader("C_CompanyID")) Then
        _C_CompanyID = String.Empty
      Else
        _C_CompanyID = CType(Reader("C_CompanyID"), String)
      End If
      If Convert.IsDBNull(Reader("C_DivisionID")) Then
        _C_DivisionID = String.Empty
      Else
        _C_DivisionID = CType(Reader("C_DivisionID"), String)
      End If
      If Convert.IsDBNull(Reader("C_OfficeID")) Then
        _C_OfficeID = String.Empty
      Else
        _C_OfficeID = CType(Reader("C_OfficeID"), String)
      End If
      If Convert.IsDBNull(Reader("C_DepartmentID")) Then
        _C_DepartmentID = String.Empty
      Else
        _C_DepartmentID = CType(Reader("C_DepartmentID"), String)
      End If
      If Convert.IsDBNull(Reader("C_DesignationID")) Then
        _C_DesignationID = String.Empty
      Else
        _C_DesignationID = CType(Reader("C_DesignationID"), String)
      End If
      _ActiveState = CType(Reader("ActiveState"), Boolean)
      If Convert.IsDBNull(Reader("C_ConfirmedOn")) Then
        _C_ConfirmedOn = String.Empty
      Else
        _C_ConfirmedOn = CType(Reader("C_ConfirmedOn"), String)
      End If
      If Convert.IsDBNull(Reader("C_ProjectSiteID")) Then
        _C_ProjectSiteID = String.Empty
      Else
        _C_ProjectSiteID = CType(Reader("C_ProjectSiteID"), String)
      End If
      If Convert.IsDBNull(Reader("VerifierID")) Then
        _VerifierIDEmployeeName = String.Empty
      Else
        _VerifierIDEmployeeName = Reader("HRM_Employees1_EmployeeName") & " [" & CType(Reader("VerifierID"), String) & "]"
      End If
      If Convert.IsDBNull(Reader("ApproverID")) Then
        _ApproverIDEmployeeName = String.Empty
      Else
        _ApproverIDEmployeeName = Reader("HRM_Employees2_EmployeeName") & " [" & CType(Reader("ApproverID"), String) & "]"
      End If
      _Contractual = CType(Reader("Contractual"), Boolean)
    End Sub
    Public Sub New()
    End Sub
  End Class
End Namespace

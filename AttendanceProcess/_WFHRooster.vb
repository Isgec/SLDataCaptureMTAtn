Imports System.Data.SqlClient
Namespace SIS.ATN
  Partial Public Class WFHRooster
    Public Property CardNo As String = ""
    Private _AttenDate As String = ""
    Public Property WFH1stHalf As Boolean = False
    Public Property WFH2ndHalf As Boolean = False
    Public Property WFHFullDay As Boolean = False
    Public Property NonWorkingDay As Boolean = False
    Public Property EmployeeNotActive As Boolean = False
    Public Property RoosterStatus As Int32 = 0
    Public Property CreatedBy As String = ""
    Public Property CreatedOn As String = ""
    Public Property CreaterRemarks As String = ""
    Public Property ModifiedBy As String = ""
    Public Property ModifiedOn As String = ""
    Public Property ModifierRemarks As String = ""
    Public Property LockedBy As String = ""
    Public Property LockedOn As String = ""
    Public Property LockerRemarks As String = ""
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
    Sub New(rd As SqlDataReader)
      SIS.SYS.SQLDatabase.DBCommon.NewObj(Me, rd)
    End Sub
    Public Sub New()
    End Sub
    Public Shared Function GetByCardDate(cardno As String, attendate As String) As SIS.ATN.WFHRooster
      Dim tmp As SIS.ATN.WFHRooster = Nothing
      Using Con As SqlConnection = New SqlConnection(SIS.SYS.SQLDatabase.DBCommon.GetConnectionString())
        Con.Open()
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.Text
          Cmd.CommandText = "select * from ATN_WFHRooster where CardNo='" & cardno & "' and attendate=convert(datetime,'" & attendate & "',103) "  'and roosterstatus=4" '4=>Approved
          Dim rd As SqlDataReader = Cmd.ExecuteReader()
          While (rd.Read())
            tmp = New SIS.ATN.WFHRooster(rd)
          End While
          rd.Close()
        End Using
      End Using
      Return tmp
    End Function
  End Class
End Namespace

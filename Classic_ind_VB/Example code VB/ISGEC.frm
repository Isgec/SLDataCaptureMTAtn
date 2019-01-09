VERSION 5.00
Begin VB.Form IsgecCode 
   BackColor       =   &H80000000&
   Caption         =   "Classic_Ind_ Sample Code"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   7545
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   285
      Left            =   5220
      TabIndex        =   3
      Top             =   720
      Width           =   765
   End
   Begin VB.Timer Tmr1 
      Left            =   3210
      Top             =   210
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   285
      Left            =   3210
      TabIndex        =   2
      Top             =   960
      Width           =   765
   End
   Begin VB.TextBox TxtDateDelete 
      Height          =   345
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   1275
   End
   Begin VB.TextBox txtCaptureDataDate 
      Height          =   345
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   1275
   End
End
Attribute VB_Name = "IsgecCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I As Integer
Dim Msg
Dim Started As Boolean
Dim mc() As String

'Device Connect
Private Declare Function device_connect Lib "Classic_Ind_MC.dll" (ByVal ipadd As String, ByVal port As Integer) As Integer
'Get Version
Private Declare Function get_version Lib "Classic_Ind_MC.dll" () As uds
'Read Template
Private Declare Function template_read Lib "Classic_Ind_MC.dll" () As uds
'Read Template
Private Declare Function get_finger Lib "Classic_Ind_MC.dll" () As uds
'date set
Private Declare Function set_date Lib "Classic_Ind_MC.dll" (ByVal STR As String) As Integer
'Restart Machine
Private Declare Function restart Lib "Classic_Ind_MC.dll" () As Integer
'Reset Machine
Private Declare Function Reset_Machine Lib "Classic_Ind_MC.dll" () As Integer
'Capture Data
Private Declare Function capture_data Lib "Classic_Ind_MC.dll" (ByVal STR As String) As uds
'Template Download
Private Declare Function bio_temp_cmd Lib "Classic_Ind_MC.dll" (ByVal STR As String) As uds
'Biometric Version
Private Declare Function bio_version Lib "Classic_Ind_MC.dll" () As uds
'date received
Private Declare Function get_date_data Lib "Classic_Ind_MC.dll" () As uds
'SET IP Address
Private Declare Function set_ipadd Lib "Classic_Ind_MC.dll" (ByVal STR As String) As udf
'Receive Card List
Private Declare Function card_list Lib "Classic_Ind_MC.dll" () As uds
' Enroll Template
Private Declare Function enroll_template Lib "Classic_Ind_MC.dll" (ByVal empname As String, ByVal empcard As String, ByVal temp As String) As udf
' write Template in card
Private Declare Function write_temp_card Lib "Classic_Ind_MC.dll" (ByVal empname As String, ByVal empcard As String, ByVal temp As String) As udf
' write Template in card
Private Declare Function write_temp_append Lib "Classic_Ind_MC.dll" (ByVal empname As String, ByVal empcard As String, ByVal temp As String) As udf
'Battery Status in volt
Private Declare Function battery_status Lib "Classic_Ind_MC.dll" () As uds
'Battery Status in Minivolt
Private Declare Function battery_status_mv Lib "Classic_Ind_MC.dll" () As Integer
'machine IN/OUT Mode
Private Declare Function get_inout_status Lib "Classic_Ind_MC.dll" () As uds
'Get_data to Get data from machine with every function
Private Declare Function get_data Lib "Classic_Ind_MC.dll" () As GetData
'device_disconnect
Private Declare Function device_close Lib "Classic_Ind_MC.dll" () As Integer
'time set
Private Declare Function set_time Lib "Classic_Ind_MC.dll" (ByVal STR As String) As Integer
'delete_finger
Private Declare Function Delete_Finger Lib "Classic_Ind_MC.dll" (ByVal STR As String) As Integer
'write_card
Private Declare Function Write_Card Lib "Classic_Ind_MC.dll" (ByVal STR As String) As Integer
'gateway address set
Private Declare Function set_gateway Lib "Classic_Ind_MC.dll" (ByVal STR As String) As udf
'netmask address setting
Private Declare Function set_netmask Lib "Classic_Ind_MC.dll" (ByVal STR As String) As udf
'delete card list
Private Declare Function delete_card_list Lib "Classic_Ind_MC.dll" () As Integer
'date delete
Private Declare Function date_delete Lib "Classic_Ind_MC.dll" (ByVal STR As String) As Integer
'get memory percentege
Private Declare Function memory_percent Lib "Classic_Ind_MC.dll" () As uds
'get date time od machine
Private Declare Function get_date_time Lib "Classic_Ind_MC.dll" () As uds
'set in_out Mode
Private Declare Function set_inout_status Lib "Classic_Ind_MC.dll" (ByVal STR As String) As Integer
'dll version
Private Declare Function dll_version Lib "Classic_Ind_MC.dll" () As udf
'send wave file
Private Declare Function send_wav_file Lib "Classic_Ind_MC.dll" (ByVal STR As String, ByVal STR1 As String) As udf

Private Type uds
    Data As String * 2048
    size As Integer
    TotalSize As String * 10
    status As Integer
End Type

Private Type GetData
    PickData As String * 2048
    size As Integer
End Type

Private Type udf
    errdata As String * 50
    err_status As Integer
End Type
Dim T_Data As String
Dim TotalCount As Double

Private Sub cmdStart_Click()
    ReDim mc(0)
    Dim fs As New FileSystemObject
    Dim tx As TextStream
    Set tx = fs.OpenTextFile(App.Path & "\IsgecMC.txt")
    Dim I As Integer
    I = 0
    Do While Not tx.AtEndOfStream
        ReDim Preserve mc(I)
        mc(I) = tx.ReadLine
        I = I + 1
    Loop
    tx.Close
    GetISGECData

'Started = True
'Tmr1.Interval = 5000
'Tmr1.Enabled = True
End Sub
Private Sub cmdStop_Click()
Started = False
End Sub

Private Sub Tmr1_Timer()
Tmr1.Enabled = False
GetISGECData
If Started = True Then
    Tmr1.Enabled = True
End If
End Sub
Private Sub GetISGECData()
    Dim J As Integer
    Dim I As Integer
    I = 0
    For I = 0 To UBound(mc)
        Dim ARY
        ARY = Split(mc(I), ",")
        J = device_connect(CStr(Trim(ARY(1))), 1085)
        If J = 0 Then 'Connected
            Dim sdt As String
            Dim tdt As String
            sdt = txtCaptureDataDate.Text
            tdt = TxtDateDelete.Text
            Do While sdt <> tdt
                GetData (sdt)
                Dim tmp As Date
                tmp = CDate(Mid(sdt, 1, 2) & "/" & Mid(sdt, 3, 2) & "/20" & Mid(sdt, 5, 2))
                tmp = DateAdd("d", 1, tmp)
                sdt = Right("00" & Day(tmp), 2) & Right("00" & Month(tmp), 2) & Right("00" & Year(tmp), 2)
            Loop
            GetData (sdt)
        End If 'Connected
        J = device_close()
    Next 'mc Loop
End Sub
Private Sub GetData(ByVal dt As String)
    Dim Struct_Info As uds
    Dim Strdata As String
    Dim Struct_Getdata As GetData
    T_Data = ""
    TotalCount = 0
    Dim I As Integer
    Struct_Info = capture_data(Trim(dt))
    If Struct_Info.status = 0 Then
        Strdata = Struct_Info.Data
        T_Data = T_Data & Strdata
        TotalCount = Struct_Info.size
        If IsNumeric(Struct_Info.TotalSize) Then
            If Struct_Info.size = CDbl(Struct_Info.TotalSize) Then
            Else
                Do While I = 0
                    Struct_Getdata = get_data()
                    TotalCount = TotalCount + Struct_Getdata.size
                    T_Data = T_Data + Struct_Getdata.PickData
                    If CDbl(Struct_Info.TotalSize) = TotalCount Then
                        I = 1
                    End If
                Loop
            End If
            UpdateSQL (T_Data)
        End If
    ElseIf Struct_Info.status = 1 Then
        'write_info ("No Data Found")
    Else
        'write_info ("Connection Problem")
    End If

End Sub
Public Function UpdateSQL(ByVal s As String)
Dim cn As New ADODB.Connection
cn.CursorLocation = adUseClient
cn.ConnectionString = "Provider=SQLOLEDB.1;Password=Webpay@2013;Persist Security Info=True;User ID=sa;Initial Catalog=ISGEC;Data Source=192.9.200.169"
cn.Open
Dim rs As ADODB.Recordset

Dim FileName As String
Dim a
Dim STR As String
Dim dat As String
Dim cardno As String
STR = "20" & Mid(s, 5, 2) & "-" & Mid(s, 3, 2) & "-" & Mid(s, 1, 2) 'yyyy-mm-dd
a = Split(s, Chr(10) & Chr(13))
For I = 1 To UBound(a, 1) - 2
    cardno = Trim(Mid(a(I), 5, 8))
    dat = STR & " " & Mid(a(I), 1, 2) & ":" & Mid(a(I), 3, 2) & ":00.000"
    Dim sql As String
    sql = "select * from trnpunchtime where fk_emp_code='" & cardno & "' and trndate='" & dat & "'"
    Set rs = cn.Execute(sql)
    If Not rs Is Nothing Then
        If rs.RecordCount <= 0 Then
        On Error Resume Next
        sql = "insert trnpunchtime (fk_emp_code,trndate) values ('" & cardno & "','" & dat & "')"
        cn.Execute sql
        End If
        Else
        MsgBox "found"
    End If
Next
cn.Close
End Function


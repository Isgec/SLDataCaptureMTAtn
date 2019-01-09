VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDevice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public TotalCount As Long
'Device Connect
Private Declare Function device_connect Lib "Classic_Ind_MC.dll" (ByVal ipadd As String, ByVal Port As Integer) As Integer
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
Public Function Connect(ByVal IPAddress As String, ByVal Port As Integer) As Integer
    Dim I As Integer
    I = device_connect(IPAddress, Port)
    Connect = I
End Function
Public Function Disconnect() As Integer
    Dim I As Integer
    I = device_close()
    Disconnect = I
End Function
Public Function GetPunchData(ByVal ddmmyy As String) As String
    Dim Struct_Info As uds
    Dim Strdata As String
    Dim Struct_Getdata As GetData
    Dim T_Data  As String
    TotalCount = 0
    Dim I As Integer
    Struct_Info = capture_data(Trim(ddmmyy))
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
            GetPunchData = T_Data
        End If
    ElseIf Struct_Info.status = 1 Then
        'write_info ("No Data Found")
    Else
        'write_info ("Connection Problem")
    End If
End Function

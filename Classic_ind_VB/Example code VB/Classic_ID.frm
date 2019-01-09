VERSION 5.00
Begin VB.Form SampleCode 
   BackColor       =   &H80000000&
   Caption         =   "Classic_Ind_ Sample Code"
   ClientHeight    =   9105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9105
   ScaleWidth      =   13875
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   285
      Left            =   390
      TabIndex        =   69
      Top             =   2970
      Width           =   765
   End
   Begin VB.Frame Frame4 
      Caption         =   "Data Display"
      Height          =   8925
      Left            =   8550
      TabIndex        =   60
      Top             =   90
      Width           =   5265
      Begin VB.TextBox txtmessage 
         Height          =   8625
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   61
         Top             =   210
         Width           =   5115
      End
   End
   Begin VB.Frame Frame3 
      Height          =   765
      Left            =   150
      TabIndex        =   58
      Top             =   8250
      Width           =   8325
      Begin VB.CommandButton CmdBtnExit 
         Caption         =   "Close"
         Height          =   375
         Left            =   3390
         TabIndex        =   59
         Top             =   270
         Width           =   1700
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Others"
      Height          =   1785
      Left            =   4500
      TabIndex        =   52
      Top             =   6360
      Width           =   3975
      Begin VB.TextBox Txtcardtowrite 
         Height          =   345
         Left            =   2070
         TabIndex        =   62
         Top             =   480
         Width           =   1275
      End
      Begin VB.CommandButton CmdRestart 
         Caption         =   "Restart Device"
         Height          =   375
         Left            =   2070
         TabIndex        =   55
         Top             =   1170
         Width           =   1700
      End
      Begin VB.CommandButton CmdReset 
         Caption         =   "Reset Device"
         Height          =   375
         Left            =   180
         TabIndex        =   54
         Top             =   1170
         Width           =   1700
      End
      Begin VB.CommandButton CMD_write_card 
         Caption         =   "Card Write"
         Height          =   375
         Left            =   180
         TabIndex        =   53
         Top             =   450
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Card List"
      Height          =   915
      Left            =   4500
      TabIndex        =   49
      Top             =   5340
      Width           =   3945
      Begin VB.CommandButton cmdReceiveCardList 
         Caption         =   "Receive Card List"
         Height          =   375
         Left            =   180
         TabIndex        =   51
         Top             =   360
         Width           =   1700
      End
      Begin VB.CommandButton CmdDelCardlist 
         Caption         =   "Delete Card List"
         Height          =   375
         Left            =   2100
         TabIndex        =   50
         Top             =   360
         Width           =   1700
      End
   End
   Begin VB.Frame Frmdata 
      Caption         =   "Data Capture"
      Height          =   1755
      Left            =   4500
      TabIndex        =   43
      Top             =   3480
      Width           =   3945
      Begin VB.CommandButton CmdDateList 
         Caption         =   "Date List "
         Height          =   375
         Left            =   1110
         TabIndex        =   48
         Top             =   240
         Width           =   1700
      End
      Begin VB.CommandButton CmdCaptureDate 
         Caption         =   "Capture Data"
         Height          =   375
         Left            =   120
         TabIndex        =   47
         Top             =   720
         Width           =   1700
      End
      Begin VB.TextBox txtCaptureDataDate 
         Height          =   345
         Left            =   2550
         TabIndex        =   46
         Top             =   750
         Width           =   1275
      End
      Begin VB.CommandButton CmdDelDate 
         Caption         =   "Delete Data"
         Height          =   375
         Left            =   120
         TabIndex        =   45
         Top             =   1290
         Width           =   1700
      End
      Begin VB.TextBox TxtDateDelete 
         Height          =   345
         Left            =   2550
         TabIndex        =   44
         Top             =   1320
         Width           =   1275
      End
      Begin VB.Label Label3 
         Caption         =   "(ddmmyy)"
         Height          =   225
         Index           =   1
         Left            =   1860
         TabIndex        =   57
         Top             =   900
         Width           =   675
      End
      Begin VB.Label Label3 
         Caption         =   "(ddmmyy)"
         Height          =   225
         Index           =   0
         Left            =   1860
         TabIndex        =   56
         Top             =   1440
         Width           =   675
      End
   End
   Begin VB.Frame Frmbasic 
      Caption         =   "Machine Status/Information"
      Height          =   1875
      Left            =   150
      TabIndex        =   39
      Top             =   1500
      Width           =   4215
      Begin VB.CommandButton cmdStart 
         Caption         =   "Start"
         Height          =   285
         Left            =   240
         TabIndex        =   68
         Top             =   1170
         Width           =   765
      End
      Begin VB.Timer Tmr1 
         Left            =   240
         Top             =   420
      End
      Begin VB.CommandButton CmdMemory 
         Caption         =   "Memory(Percentage)"
         Height          =   375
         Left            =   1320
         TabIndex        =   42
         Top             =   1350
         Width           =   1700
      End
      Begin VB.CommandButton CmdDateTime 
         Caption         =   "Date&&Time"
         Height          =   375
         Left            =   1320
         TabIndex        =   41
         Top             =   840
         Width           =   1700
      End
      Begin VB.CommandButton CmdGetVersion 
         Caption         =   "Version"
         Height          =   375
         Left            =   1320
         TabIndex        =   40
         Top             =   330
         Width           =   1700
      End
   End
   Begin VB.Frame Frmnetset 
      Caption         =   "Network Setting"
      Height          =   1875
      Left            =   4470
      TabIndex        =   29
      Top             =   1500
      Width           =   3975
      Begin VB.TextBox TxtsetIPAddress 
         Height          =   345
         Left            =   1470
         TabIndex        =   38
         Top             =   330
         Width           =   1275
      End
      Begin VB.CommandButton CmdIPAddress 
         Caption         =   "Apply"
         Height          =   375
         Left            =   3030
         TabIndex        =   37
         Top             =   300
         Width           =   700
      End
      Begin VB.CommandButton CmdSetgw 
         Caption         =   "Apply"
         Height          =   375
         Left            =   3030
         TabIndex        =   33
         Top             =   1290
         Width           =   700
      End
      Begin VB.TextBox TxtSetGateway 
         Height          =   345
         Left            =   1470
         TabIndex        =   32
         Top             =   1320
         Width           =   1275
      End
      Begin VB.CommandButton CmdSetNetmask 
         Caption         =   "Apply"
         Height          =   375
         Left            =   3030
         TabIndex        =   31
         Top             =   810
         Width           =   700
      End
      Begin VB.TextBox TxtSetNetmask 
         Height          =   345
         Left            =   1470
         TabIndex        =   30
         Top             =   840
         Width           =   1275
      End
      Begin VB.Label Label8 
         Caption         =   "IP Address"
         Height          =   255
         Left            =   270
         TabIndex        =   36
         Top             =   450
         Width           =   1155
      End
      Begin VB.Label Label7 
         Caption         =   "Gateway"
         Height          =   315
         Index           =   1
         Left            =   270
         TabIndex        =   35
         Top             =   1380
         Width           =   1005
      End
      Begin VB.Label Label7 
         Caption         =   "Subnet Mask"
         Height          =   255
         Index           =   0
         Left            =   270
         TabIndex        =   34
         Top             =   900
         Width           =   1245
      End
   End
   Begin VB.Frame Frmdatetime 
      Caption         =   "Date && Time Setting"
      Height          =   1300
      Left            =   4470
      TabIndex        =   21
      Top             =   90
      Width           =   3975
      Begin VB.CommandButton CmdSetTime 
         Caption         =   "Apply"
         Height          =   375
         Left            =   3030
         TabIndex        =   25
         Top             =   810
         Width           =   700
      End
      Begin VB.TextBox txtTimeSet 
         Height          =   345
         Left            =   1470
         TabIndex        =   24
         Top             =   810
         Width           =   1275
      End
      Begin VB.CommandButton CmdDateSet 
         Caption         =   "Apply"
         Height          =   375
         Left            =   3030
         TabIndex        =   23
         Top             =   270
         Width           =   700
      End
      Begin VB.TextBox txtDateSet 
         Height          =   345
         Left            =   1470
         TabIndex        =   22
         Top             =   270
         Width           =   1275
      End
      Begin VB.Label Label6 
         Caption         =   "Time(hhmmss)"
         Height          =   345
         Left            =   240
         TabIndex        =   27
         Top             =   870
         Width           =   1245
      End
      Begin VB.Label Label5 
         Caption         =   "Date(ddmmyy)"
         Height          =   315
         Left            =   240
         TabIndex        =   26
         Top             =   300
         Width           =   1245
      End
   End
   Begin VB.Frame Frmbiometric 
      Caption         =   "Bimetric"
      Height          =   4695
      Left            =   150
      TabIndex        =   7
      Top             =   3480
      Width           =   4215
      Begin VB.TextBox Txtaddfinger 
         Height          =   345
         Left            =   2400
         TabIndex        =   20
         Top             =   4170
         Width           =   1275
      End
      Begin VB.CommandButton CmdwriteFinger_card 
         Caption         =   "Write Card with finger"
         Height          =   375
         Left            =   330
         TabIndex        =   19
         Top             =   3510
         Width           =   1700
      End
      Begin VB.TextBox txtTemp 
         Height          =   375
         Left            =   2400
         TabIndex        =   18
         Top             =   3480
         Width           =   1275
      End
      Begin VB.CommandButton Cmdappendfinger 
         Caption         =   "Add Finger in Card"
         Height          =   375
         Left            =   330
         TabIndex        =   17
         Top             =   4140
         Width           =   1700
      End
      Begin VB.TextBox txtcardnum 
         Height          =   345
         Left            =   2400
         TabIndex        =   16
         Top             =   900
         Width           =   1275
      End
      Begin VB.CommandButton Cmdgetfinger 
         Caption         =   "Enroll New Finger"
         Height          =   375
         Left            =   330
         TabIndex        =   15
         Top             =   930
         Width           =   1700
      End
      Begin VB.CommandButton CmdDownloadTemplate 
         Caption         =   "Finger Download"
         Height          =   375
         Left            =   330
         TabIndex        =   14
         Top             =   1560
         Width           =   1700
      End
      Begin VB.TextBox txtCardNo 
         Height          =   345
         Left            =   2400
         TabIndex        =   13
         Top             =   1590
         Width           =   1275
      End
      Begin VB.CommandButton CmdEnrollTemplate 
         Caption         =   "Finger Upload"
         Height          =   375
         Left            =   330
         TabIndex        =   12
         Top             =   2220
         Width           =   1700
      End
      Begin VB.TextBox txtTempEnroll 
         Height          =   345
         Left            =   2400
         TabIndex        =   11
         Top             =   2250
         Width           =   1275
      End
      Begin VB.CommandButton CmdDel_fin 
         Caption         =   "Delete Finger"
         Height          =   375
         Left            =   330
         TabIndex        =   10
         Top             =   2880
         Width           =   1700
      End
      Begin VB.TextBox TxtCARD 
         Height          =   345
         Left            =   2400
         TabIndex        =   9
         Top             =   2880
         Width           =   1275
      End
      Begin VB.CommandButton BioVersion 
         Caption         =   "Bio Version"
         Height          =   375
         Left            =   330
         TabIndex        =   8
         Top             =   300
         Width           =   1700
      End
      Begin VB.Label Label4 
         Caption         =   "Card No."
         Height          =   195
         Index           =   5
         Left            =   2400
         TabIndex        =   67
         Top             =   3990
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Card No."
         Height          =   195
         Index           =   4
         Left            =   2400
         TabIndex        =   66
         Top             =   3300
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Card No."
         Height          =   195
         Index           =   3
         Left            =   2400
         TabIndex        =   65
         Top             =   2700
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Card No."
         Height          =   195
         Index           =   2
         Left            =   2400
         TabIndex        =   64
         Top             =   1410
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "File Name"
         Height          =   195
         Index           =   1
         Left            =   2400
         TabIndex        =   63
         Top             =   2070
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Card No."
         Height          =   195
         Index           =   0
         Left            =   2400
         TabIndex        =   28
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.Frame Frmconnect 
      Caption         =   "Network Connection"
      Height          =   1300
      Left            =   150
      TabIndex        =   0
      Top             =   90
      Width           =   4215
      Begin VB.CommandButton CmdDisconnect 
         Caption         =   "Disconnect"
         Height          =   340
         Left            =   2610
         TabIndex        =   6
         Top             =   750
         Width           =   1515
      End
      Begin VB.TextBox txtPortNo 
         Height          =   345
         Left            =   1200
         TabIndex        =   3
         Text            =   "1085"
         Top             =   720
         Width           =   1275
      End
      Begin VB.TextBox TxtIPAddress 
         Height          =   345
         Left            =   1200
         TabIndex        =   2
         Text            =   "192.168.0.88"
         Top             =   270
         Width           =   1275
      End
      Begin VB.CommandButton CmdConnect 
         Caption         =   "Connect "
         Height          =   375
         Left            =   2610
         TabIndex        =   1
         Top             =   300
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "Port No"
         Height          =   225
         Left            =   120
         TabIndex        =   5
         Top             =   750
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "IP Address"
         Height          =   225
         Left            =   120
         TabIndex        =   4
         Top             =   330
         Width           =   1095
      End
   End
End
Attribute VB_Name = "SampleCode"
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
'This function returns the biometric version of machine
'return structure has status of success/fail, count of return bytes also.
Private Sub BioVersion_Click()
    Dim Struct_Info As uds
    Dim Data As String
    Dim i1 As Integer
    Struct_Info = bio_version()
    If Struct_Info.status = 0 Then
        Data = Struct_Info.Data
        write_info (Data)
        write_info ("Size of Data is -:" & Struct_Info.size)
    Else
        write_info ("Version Not Getted")
        Exit Sub
    End If
End Sub
'To write a number in card
Private Sub CMD_write_card_Click()
I = Write_Card(Trim(txtcardnum.Text))
    If I = 0 Then
        write_info ("Done Successfully")
    Else
        write_info ("Check Card Number, Not Changed Successfully")
    End If
End Sub
'To close application
Private Sub CmdBtnExit_Click()
End
End Sub
'This function returns the punching data of given date from machine
'return structure has status of success/fail, count of return bytes also.
Private Sub CmdCaptureDate_Click()
    Dim Struct_Info As uds
    Dim Strdata As String
    Dim Struct_Getdata As GetData
    T_Data = ""
    TotalCount = 0
    Dim I As Integer
    Struct_Info = capture_data(Trim(txtCaptureDataDate.Text))
    If Struct_Info.status = 0 Then
        Strdata = Struct_Info.Data
        T_Data = T_Data & Strdata
        TotalCount = Struct_Info.size
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
        Write_File_data (T_Data)
        write_info (T_Data)
        write_info ("Size of Data is -:" & Struct_Info.TotalSize)
    ElseIf Struct_Info.status = 1 Then
        write_info ("No Data Found")
    Else
        write_info ("Connection Problem")
    End If

End Sub

'This function Set the Time of machine
'return structure has status of success/ fail.
Private Sub CmdConnect_Click()
    I = device_connect(CStr(Trim(TxtIPAddress.Text)), CInt(Trim(txtPortNo.Text)))
    If I = 0 Then
        MsgBox "Device Successfully Connected", vbInformation
    Else
        MsgBox " Device Not Connected", vbCritical
    End If
End Sub

'To display return data in edit box
Public Function write_info(ByVal s As String)
        s = Replace(s, Chr(127), "")
        s = Replace(s, Chr(13), "")
        If InStr(1, s, Chr(10)) Then
            Msg = Split(s, Chr(10))
            For I = 0 To UBound(Msg)
                txtmessage.SelStart = Len(txtmessage.Text)
                txtmessage.SelText = Msg(I) & vbCrLf
                txtmessage.Refresh
                'txtmessage.Text = txtmessage.Text & Msg(i) & vbCrLf
            Next
    Else
        txtmessage.SelStart = Len(txtmessage.Text)
        txtmessage.SelText = s & vbCrLf
        txtmessage.Refresh
        'txtmessage.Text = txtmessage.Text & s & vbCrLf & vbCrLf
    End If
End Function

'This function returns the total date list of punching date, which is save in machine
'return structure has status of success/fail, count of return bytes also.
Private Sub CmdDateList_Click()
    Dim Struct_Info As uds
    Dim Strdata As String
    Dim Struct_Getdata As GetData
    T_Data = ""
    TotalCount = 0
    Dim I As Integer
    Struct_Info = get_date_data()
    If Struct_Info.status = 0 Then
        Strdata = Struct_Info.Data
        T_Data = T_Data & Strdata
        TotalCount = Struct_Info.size
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
        write_info (T_Data)
        write_info ("Size of Data is -:" & Struct_Info.TotalSize)
    ElseIf Struct_Info.status = 1 Then
        write_info ("No Data Found")
    Else
        write_info ("Connection Problem")
    End If
End Sub

'This function Set the Time of machine
'return structure has status of success/fail.
Private Sub CmdDateSet_Click()
    I = set_date(Trim(txtDateSet.Text))
    If I = 0 Then
        write_info ("Date is Changed Successfully")
    Else
        write_info ("Check Date Format, Not Changed Successfully")
    End If
End Sub

'This function returns the date and time of machine
'return structure has status of success/fail, number of return bytes.
Private Sub CmdDateTime_Click()
    Dim Struct_Info As uds
    Dim Strdata As String
    Struct_Info = get_date_time()
    If Struct_Info.status = 0 Then
        Strdata = Struct_Info.Data
        write_info (Strdata)
        write_info ("Size of Data is :-" & Struct_Info.size)
    Else
        write_info ("Connection Problem")
    End If
End Sub

'This function Delete the finger which is include in machine
'return structure has status of success/fail.
Private Sub CmdDel_fin_Click()
    I = Delete_Finger(TxtCARD.Text)     ' card number must be with eight digit
    If I = 0 Then
        write_info ("Deleted Successfully")
    Else
        write_info ("Check time Format, Not Changed Successfully")
    End If
End Sub

'This function Delete the card list which is include in machine
'return structure has status of success/fail.
Private Sub CmdDelCardlist_Click()
    I = delete_card_list()
    If I = 0 Then
        write_info ("Card list deleted Successfully")
    Else
        write_info ("Card list not deleted, Check Your Connection")
    End If
End Sub

'This function delete the punching data of given date from machine
'return structure has status of success/fail.
Private Sub CmdDelDate_Click()
    I = date_delete(Trim(TxtDateDelete.Text))
    If I = 0 Then
        write_info ("Date Data deleted Successfully")
    Else
        write_info ("date data not deleted, Check Your Connection")
    End If
End Sub

' to Disconnect the machine
Private Sub CmdDisconnect_Click()
    I = device_close()
    If I = 0 Then
        write_info ("Device Disconnect Successfully")
    Else
        write_info ("Device is not Closed, Check Your Connection")
    End If
End Sub

'Show the version of dll......
Private Sub CmdDllVer_Click()
Dim get_error As udf
Dim Data As String
get_error = dll_version()
    Data = get_error.errdata
    write_info (Data)

End Sub

'This function returns the finger data of given card number, which is save in machine
'return structure has status of success/ fail, count of return bytes also.Private Sub CmdDownloadTemplate_Click()
Private Sub CmdDownloadTemplate_Click()
    Dim Struct_Info As uds
    Dim Data As String
    Dim i1 As Integer
    Struct_Info = bio_temp_cmd(Trim(txtCardNo.Text))
    If Struct_Info.status = 0 Then
        Data = Struct_Info.Data
        write_info (Data)
        save_finger_file (Data)
        write_info ("Size of Data is -:" & Struct_Info.size)
    ElseIf Struct_Info.status = 1 Then
        write_info ("No Template Found")
    Else
        write_info ("Error")
    End If

End Sub

' to send finger data to machine
'This function returns the error
'return structure has status of success/fail.
Private Sub CmdEnrollTemplate_Click()
Dim get_error As udf
Dim Data As String
Dim lngImageSize As Long
Dim lngOffset As Long
Dim bytChunk() As Byte
Dim strTempPic As String
Dim fs As New FileSystemObject
Dim tx As TextStream
Dim intFile
intFile = FreeFile()
Dim mFile_Name As String, FNum As Integer
Dim astr As String, LogStr As String
strTempPic = App.Path & "\" & Trim(txtTempEnroll)
LogStr = ""
Open strTempPic For Binary As #intFile
Set tx = fs.OpenTextFile(strTempPic, ForReading, True)
 
    LogStr = tx.Read(385)
    tx.Close

    get_error = enroll_template("anuj choudhary", "00000007", LogStr)
    If get_error.err_status = 0 Then
        Data = get_error.errdata
        write_info (Data)
    ElseIf get_error.err_status = 1 Then
        Data = get_error.errdata
        write_info (Data)
    Else
        write_info ("Check your Network Connection")
    End If
End Sub

'This function read the finger data which, puts at executing time of this command at the biometric sencer of machine
' it returns status of success/fail, and error too, if any....
Private Sub Cmdgetfinger_Click()
    Dim Struct_Info As uds
    Dim Strdata As String
    Dim k
    Struct_Info = template_read()
   If Struct_Info.status = 0 Then
        Strdata = Struct_Info.Data
        write_info (Strdata)
        save_file (Struct_Info.Data)
        write_info ("Size of Data is -:" & Struct_Info.size)
    ElseIf Struct_Info.status = 1 Then
        write_info (Struct_Info.Data)
    Else
        write_info ("Connection Problem" & vbCrLf)
    End If
    
End Sub

'This function returns the version of machine
'return structure has status of success/ fail, count of return bytes also.
Private Sub CmdGetVersion_Click()
    Dim Struct_Info As uds
    Dim Data As String
    Dim i1 As Integer
    Struct_Info = get_version()
    If Struct_Info.status = 0 Then
        Data = Struct_Info.Data
        write_info (Data)
        write_info ("Size of Data is -:" & Struct_Info.size)
    Else
        write_info ("Version Not Getted")
        Exit Sub
    End If
End Sub

'This function Set the IP Address of machine
'return structure has status of success/ fail, error type also.
Private Sub CmdIPAddress_Click()
    Dim get_error As udf
    Dim Strdata As String
    get_error = set_ipadd(Trim(TxtsetIPAddress.Text))
    If get_error.err_status = 0 Then
        write_info ("IP Address Changes Successfully, Restart Machine")
    ElseIf get_error.err_status = 1 Then
        Strdata = get_error.errdata
        write_info (Strdata)
    Else
        write_info ("IP Change Failure, Check Your Connection")
    End If
End Sub

'This function returns the Memory percentage of machine
'return structure has status of success/ fail,.
Private Sub CmdMemory_Click()
    Dim Struct_Info As uds
    Dim Strdata As String
    Struct_Info = memory_percent()
    If Struct_Info.status = 0 Then
        Strdata = Struct_Info.Data
        write_info (Strdata)
        write_info ("Size of Data is -:" & Struct_Info.size)
    ElseIf Struct_Info.status = 1 Then
        write_info ("Card List Is Empty")
    Else
        write_info ("Connection Problem" & vbCrLf)
    End If
End Sub

'This function returns the card list of machine
'return structure has status of success/ fail, count of return bytes also.
Private Sub cmdReceiveCardList_Click()
   Dim Struct_Info As uds
    Dim Strdata As String
    Dim Struct_Getdata As GetData
    T_Data = ""
    TotalCount = 0
    Dim I As Integer
    Struct_Info = card_list()
    If Struct_Info.status = 0 Then
        Strdata = Struct_Info.Data
        T_Data = T_Data & Strdata
        TotalCount = Struct_Info.size
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
        write_info (T_Data)
        save_card_list (T_Data)
        write_info ("Size of Data is -:" & Struct_Info.TotalSize)
    ElseIf Struct_Info.status = 1 Then
        write_info ("No Data Found")
    Else
        write_info ("Connection Problem")
    End If
End Sub
'To reset all values
Private Sub CmdReset_Click()
    I = Reset_Machine()
    If I = 0 Then
        write_info ("Device Reset Successfully")
    Else
        write_info ("Device is not Reset , Check Your Connection")
    End If
End Sub

'This function restart machine
'return structure has status of success/ fail, .
Private Sub CmdRestart_Click()
    I = restart()
    If I = 0 Then
        write_info ("Device Restarted Successfully")
    Else
        write_info ("Device is not Restarted , Check Your Connection")
    End If
End Sub

'This function Set the GateWay Address of machine
'return structure has status of success/ fail, error type also.
Private Sub CmdSetgw_Click()
    Dim get_error As udf
    Dim Strdata As String
    get_error = set_gateway(Trim(TxtSetGateway.Text))
    If get_error.err_status = 0 Then
        write_info ("Gateway Changes Successfully, Restart Machine")
    ElseIf get_error.err_status = 1 Then
        Strdata = get_error.errdata
        write_info (Strdata)
    Else
        write_info ("Change Failure, Check Your Connection")
    End If
End Sub

'This function Set the Natmask Address of machine of machine
'return structure has status of success/ fail, error type also.
Private Sub CmdSetNetmask_Click()
    Dim get_error As udf
    Dim Strdata As String
    get_error = set_netmask(Trim(TxtSetNetmask.Text))
    If get_error.err_status = 0 Then
        write_info ("SubnetMask Changes Successfully, Restart Machine")
    ElseIf get_error.err_status = 1 Then
        Strdata = get_error.errdata
        write_info (Strdata)
    Else
        write_info ("Change Failure, Check Your Connection")
    End If
End Sub

'This function Set the Time of machine of machine
'return structure has status of success/ fail.
Private Sub CmdSetTime_Click()
    I = set_time(Trim(txtTimeSet.Text))
    If I = 0 Then
        write_info ("time is Changed Successfully")
    Else
        write_info ("Check time Format, Not Changed Successfully")
    End If
End Sub

Public Function save_file(ByVal buff As String)
Dim FileName As String
Dim STR As String
Dim FNum
FNum = FreeFile()
STR = txtcardnum.Text & ".DAT"
Open App.Path & "\" & STR For Output As #FNum
    Print #FNum, buff
Close #FNum
End Function

Public Function save_finger_file(ByVal buff As String)
Dim FileName As String
Dim STR As String
Dim FNum
FNum = FreeFile()
STR = txtCardNo.Text & ".txt"
Open App.Path & "\" & STR For Output As #FNum
    Print #FNum, buff
Close #FNum
End Function
Public Function Write_File_data(ByVal s As String)
Dim FileName As String
Dim a
Dim STR As String
Dim FNum
Dim dat As String
Dim cardno As String
FNum = FreeFile()
dat = "20" & Mid(s, 5, 2) & "-" & Mid(s, 3, 2) & "-" & Mid(s, 1, 2) 'yyyy-mm-dd
STR = dat & ".Txt"
a = Split(s, Chr(10) & Chr(13))
Open App.Path & "\" & STR For Output As #FNum
For I = 1 To UBound(a, 1) - 2
    cardno = Mid(a(I), 5, 8)
    Print #FNum, "" & cardno & " " & Mid(a(I), 1, 2) & ":" & Mid(a(I), 3, 2) & " " & Mid(a(I), 13, 1)
Next
Print #FNum, a(I)
Close #FNum
End Function

Public Function save_card_list(ByVal buff As String)
Dim FileName As String
Dim a
Dim STR As String
Dim m As Integer
Dim FNum
FNum = FreeFile()
STR = "cardlist.Txt"
a = Split(buff, Chr(10) & Chr(13))
Open App.Path & "\" & STR For Output As #FNum
For m = 0 To UBound(a)
    Print #FNum, a(m)
Next
    Close #FNum
End Function
Private Sub CmdwriteFinger_card_Click()
Dim get_error As udf
Dim Data As String
Dim lngImageSize As Long
Dim lngOffset As Long
Dim bytChunk() As Byte
Dim strTempPic As String
Dim fs As New FileSystemObject
Dim tx As TextStream
Dim intFile
intFile = FreeFile()
Dim mFile_Name As String, FNum As Integer
Dim astr As String, LogStr As String
strTempPic = App.Path & "\" & Trim(txtTemp)
LogStr = ""
Open strTempPic For Binary As #intFile
Set tx = fs.OpenTextFile(strTempPic, ForReading, True)
 
    LogStr = tx.Read(385)
    tx.Close

    get_error = write_temp_card("anuj choudhary", "00000007", LogStr)
    If get_error.err_status = 0 Then
        Data = get_error.errdata
        write_info (Data)
    ElseIf get_error.err_status = 1 Then
        Data = get_error.errdata
        write_info (Data)
    Else
        write_info ("Check your Network Connection")
    End If
End Sub

Private Sub Cmdappendfinger_Click()
Dim get_error As udf
Dim Data As String
Dim lngImageSize As Long
Dim lngOffset As Long
Dim bytChunk() As Byte
Dim strTempPic As String
Dim fs As New FileSystemObject
Dim tx As TextStream
Dim intFile
intFile = FreeFile()
Dim mFile_Name As String, FNum As Integer
Dim astr As String, LogStr As String
strTempPic = App.Path & "\" & Trim(Txtaddfinger.Text)
LogStr = ""
Open strTempPic For Binary As #intFile
Set tx = fs.OpenTextFile(strTempPic, ForReading, True)
 
    LogStr = tx.Read(385)
    tx.Close

    get_error = write_temp_append("anuj choudhary", "00000007", LogStr)
    If get_error.err_status = 0 Then
        Data = get_error.errdata
        write_info (Data)
    ElseIf get_error.err_status = 1 Then
        Data = get_error.errdata
        write_info (Data)
    Else
        write_info ("Check your Network Connection")
    End If

End Sub
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
Started = True
Tmr1.Interval = 5000
Tmr1.Enabled = True
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
        J = device_connect(CStr(Trim(ARY(1))), CInt(Trim(txtPortNo.Text)))
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
        Write_File_data (T_Data)
        write_info (T_Data)
        write_info ("Size of Data is -:" & Struct_Info.TotalSize)
    ElseIf Struct_Info.status = 1 Then
        write_info ("No Data Found")
    Else
        write_info ("Connection Problem")
    End If

End Sub

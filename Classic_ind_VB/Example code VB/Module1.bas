Attribute VB_Name = "Module1"
'Private Declare Function win_Socket_Dll Lib "win_Socket_dll.dll" Alias "winsocket" (ByRef ip As String, ByRef port As String) As Integer          ' Note that if you declare the lpData parameter as String, you must pass it By Value.
'Dim pTLIApplication As New TLI.TLIApplication
'Dim pTypeLibInfo As TLI.TypeLibInfo
'Public Declare Function device_connect Lib "new_test.dll" (ByVal ipadd As String, ByVal port As Integer) As Integer
'Public Declare Function get_version Lib "new_test.dll" () As md
'Public Type md
'    data As Variant
'    Size As Integer
'    status As Integer
'End Type
Sub main()
    
    
'
'    Set pTypeLibInfo = pTLIApplication.TypeLibInfoFromFile("C:\ab\new_test.dll")
'    pTypeLibInfo.Register
    FrmBioStarSampleCode.Show
End Sub

Module modMain
  Public AutoStart As Boolean = False
  Public AutoProcessLastDate As Boolean = False
  Public AutoProcess As Boolean = False
  Sub main()
    Dim oFrm As FrmMain = Nothing
    If My.Application.CommandLineArgs.Count > 0 Then
      If My.Application.CommandLineArgs(0).ToLower = "r" Then
        AutoStart = True
      ElseIf My.Application.CommandLineArgs(0).ToLower = "l" Then
        AutoProcessLastDate = True
      ElseIf My.Application.CommandLineArgs(0).ToLower = "p" Then
        AutoProcess = True
      End If
    End If
    oFrm = New FrmMain
    Application.Run(oFrm)
    oFrm = Nothing
  End Sub
End Module

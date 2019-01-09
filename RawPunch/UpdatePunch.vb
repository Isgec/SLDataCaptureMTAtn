Public Class UpdatePunch
  Public Shared FileUnderProcess As String = ""
  Public Shared FinYear As String = ""
  Public Shared Property ConnectionString As String
    Get
      Return SIS.SYS.SQLDatabase.DBCommon.conString
    End Get
    Set(value As String)
      SIS.SYS.SQLDatabase.DBCommon.conString = value
    End Set
  End Property
  Public Shared Sub DeleteAllPunchForDate(ByVal DataDate As String)
    'Clean All the Data For Date
    SIS.ATN.atnRawPunch.DeleteRawPunchByDate(DataDate)
  End Sub
  Public Shared Sub UpdateRawData(ByVal _CardNo As String, ByVal _PunchTime As String, ByVal DataDate As DateTime)

    Dim ReplacedCardNo As String = _CardNo
    Dim CardNo As String = ReplacedCardNo

    'Get Card Replacment
    Dim oRepl As SIS.ATN.atnCardReplacement = SIS.ATN.atnCardReplacement.GetByID(ReplacedCardNo)
    If Not oRepl Is Nothing Then
      CardNo = oRepl.CardNo
    End If
    'End of Card Replacement

    Dim strPunchTime As String = ""
    If strPunchTime = "" Then
      strPunchTime = _PunchTime
    End If
    If strPunchTime = "" Then
      strPunchTime = "00:00"
    End If

    Dim PunchTime As Decimal = Convert.ToDecimal(strPunchTime.Replace(":", "."))

    Dim oRawPunch As SIS.ATN.atnRawPunch = SIS.ATN.atnRawPunch.GetRawPunchByCardNoDate(CardNo, DataDate)
    Dim Found As Boolean = True
    If oRawPunch Is Nothing Then
      oRawPunch = New SIS.ATN.atnRawPunch
      With oRawPunch
        .CardNo = CardNo
        .PunchDate = DataDate
        .FinYear = FinYear
        .Processed = False
      End With
      Found = False
    End If

    With oRawPunch
      If .Punch1Time = PunchTime Or .Punch2Time = PunchTime Then
      Else
        If .Punch1Time <= 0 Then
          .Punch1Time = PunchTime
          .FirstPunchMachine = FileUnderProcess
        Else
          If .Punch2Time <= 0 Then
            If PunchTime >= .Punch1Time Then
              .Punch2Time = PunchTime
              .SecondPunchMachine = FileUnderProcess
            Else
              .Punch2Time = .Punch1Time
              .SecondPunchMachine = .FirstPunchMachine
              .Punch1Time = PunchTime
              .FirstPunchMachine = FileUnderProcess
            End If
          Else
            If PunchTime > .Punch2Time Then
              .Punch2Time = PunchTime
              .SecondPunchMachine = FileUnderProcess
            Else
              If PunchTime < .Punch1Time Then
                .Punch1Time = PunchTime
                .FirstPunchMachine = FileUnderProcess
              End If
            End If
          End If
        End If
      End If
    End With

    If Found Then
      SIS.ATN.atnRawPunch.Update(oRawPunch)
    Else
      SIS.ATN.atnRawPunch.Insert(oRawPunch)
    End If
  End Sub

End Class

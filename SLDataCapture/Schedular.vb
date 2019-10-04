Public Class Schedular
  Inherits TimerSupport
  Implements IDisposable

  Public Property frm As FrmMain = Nothing
  Public Event SchedularStarted()
  Public Event SchedularStopped()
  Private ci As System.Globalization.CultureInfo = New System.Globalization.CultureInfo("en-GB", True)

  Private Class mySchedule
    Public Property IntervalHours As Integer = 1
    Public Property LastTime As DateTime = Now.AddHours(-1 * IntervalHours).AddMinutes(-1 * Now.Minute)
    Public Property LastDate As DateTime = Now
  End Class
  Dim schd As mySchedule = Nothing
  Public Overrides Sub Process()
    If schd Is Nothing Then
      schd = New mySchedule
    End If
    If Not frm.Processing Then
      Dim x As DateTime = Now
      If DateAndTime.DateValue(x.Date.ToString) > DateAndTime.DateValue(schd.LastDate.Date.ToString) Then
        '======================
        'Download for Last date
        '======================
        frm.radioPunch.Checked = True
        frm.F_sdt.Value = schd.LastDate
        frm.F_tdt.Value = schd.LastDate
        If Not IsStopping Then
          frm.cmdStart.PerformClick()
        End If
        '=====================
        'Process for Last Date
        '=====================
        frm.radioProcess.Checked = True
        frm.cmdStart.PerformClick()
        schd.LastDate = x
        schd.LastTime = x
        '=========================
        'download for Current Time
        '=========================
        frm.radioPunch.Checked = True
        frm.F_sdt.Value = schd.LastDate
        frm.F_tdt.Value = schd.LastDate
        If Not IsStopping Then
          frm.cmdStart.PerformClick()
        End If
      End If
      If x.Hour <> schd.LastTime.Hour Then
        schd.LastTime = x
        '=========================
        'download for Current Time
        '=========================
        frm.radioPunch.Checked = True
        frm.F_sdt.Value = schd.LastDate
        frm.F_tdt.Value = schd.LastDate
        If Not IsStopping Then
          frm.cmdStart.PerformClick()
        End If
      End If
    End If
  End Sub

  Public Overrides Sub Started()
    RaiseEvent SchedularStarted()
  End Sub

  Public Overrides Sub Stopped()
    schd = Nothing
    RaiseEvent SchedularStopped()
  End Sub

#Region "IDisposable Support"
  Private disposedValue As Boolean ' To detect redundant calls

  ' IDisposable
  Protected Overridable Sub Dispose(disposing As Boolean)
    If Not disposedValue Then
      If disposing Then
        ' TODO: dispose managed state (managed objects).
      End If
      frm = Nothing
      ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
      ' TODO: set large fields to null.
    End If
    disposedValue = True
  End Sub

  ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
  'Protected Overrides Sub Finalize()
  '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
  '    Dispose(False)
  '    MyBase.Finalize()
  'End Sub

  ' This code added by Visual Basic to correctly implement the disposable pattern.
  Public Sub Dispose() Implements IDisposable.Dispose
    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
    Dispose(True)
    ' TODO: uncomment the following line if Finalize() is overridden above.
    ' GC.SuppressFinalize(Me)
  End Sub
#End Region
End Class

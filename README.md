# Python Market Analysis
# This code was turned into an application using PowerShell. It was then transferred to a different Excel file, where the following macros were applied, and the necessary analyses were conducted through Excel.


# Refesh

# Private Declare PtrSafe Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long


Dim NextRefresh As Date

Sub StartTimer()
    ' Set the interval (10 seconds)
    NextRefresh = Now + TimeValue("00:00:10")
    Application.OnTime NextRefresh, "RefreshData"
End Sub

Sub RefreshData()
    ' Refresh all data connections
    ThisWorkbook.RefreshAll
    
    ' Schedule the next refresh
    StartTimer
End Sub

Sub StopTimer()
    On Error Resume Next
    Application.OnTime NextRefresh, "RefreshData", , False
End Sub



# Notification

Private Declare PtrSafe Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Range("H27:H29")) Is Nothing Then
        If Application.WorksheetFunction.Min(Range("H27:H29")) < 0.3 Then
            sndPlaySound "C:\Windows\Media\click.wav", &H1
            MsgBox "Dikkat: H27:H29 aralığındaki değerlerden biri 0.3'ten küçük!"
        End If
    End If
    
    If Not Intersect(Target, Range("G2:G12")) Is Nothing Then
        If Application.WorksheetFunction.Min(Range("G2:G12")) < 0.77 Then
            sndPlaySound "C:\Windows\Media\click.wav", &H1
            MsgBox "Dikkat: G2:G12 aralığındaki değerlerden biri 0.77'ten küçük!"
        End If
    End If
End Sub


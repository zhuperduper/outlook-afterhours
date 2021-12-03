Attribute VB_Name = "Module1"

Public sendDate As Date
Public sendDateTemp As Date
Public UserDeferOption As Integer
Public SendNow As String

Dim todayDate As Date
Dim wkDay As Integer
Dim sendHour As Integer


Sub SetDate()

todayDate = Now()

'Evaluate day of week, hour, and trim time from todayDate
wkDay = Weekday(todayDate, vbMonday)
sendHour = Hour(todayDate)
todayDate = DateValue(todayDate)

'Check if Sat/Sun
'Send at Monday at 8am
If wkDay >= 6 Then
sendDate = todayDate
sendDate = DateAdd("d", 8 - wkDay, sendDate)
sendDate = DateAdd("h", 8, sendDate)

Else
    'Check if before 7am
    'Send same day at 8am
    If sendHour < 7 Then
    sendDate = todayDate
    sendDate = DateAdd("h", 8, sendDate)
    
    Else
        'Check if after 6pm
        'Send next business day at 8am
        If sendHour >= 18 Then
        sendDate = todayDate
        sendDate = DateAdd("h", 8, sendDate)
        sendDate = DateAdd("d", 1, sendDate)
        
            If wkDay = 5 Then
            sendDate = DateAdd("d", 2, sendDate)
            End If
        
        Else
        'Not weekend, not early hours, not after hours
        SendNow = "Y"
        End If
    End If
End If

End Sub

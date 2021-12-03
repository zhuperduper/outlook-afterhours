VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim obj As Object
Dim Mail As Outlook.MailItem
Dim wkDay As Integer
Dim MinNow As Integer
Dim sendHour As Integer

Function getActiveMessage() As Outlook.MailItem
Dim insp As Outlook.Inspector
If TypeOf Application.ActiveWindow Is Outlook.Inspector Then
        Set insp = Application.ActiveWindow
    End If
If insp Is Nothing Then
        Dim inline As Object
        Set inline = Application.ActiveExplorer.ActiveInlineResponse
        If inline Is Nothing Then Exit Function
Set getActiveMessage = inline
    Else
       Set insp = Application.ActiveInspector
       If insp.CurrentItem.Class = olMail Then
          Set getActiveMessage = insp.CurrentItem
       Else
         Exit Function
       End If
End If
End Function
Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
'On Error GoTo ErrorHandler

'Emails are held by default
SendNow = "N"

'Sets the sendDate and determines if email should be sent now
Call SetDate

UserDeferOption = 0

'Send the Email
Set obj = getActiveMessage()
If obj Is Nothing Then
    'Do nothing - as this is likely a calendar issue
    'MsgBox "No active inspector"
Else
    If TypeOf obj Is Outlook.MailItem Then
        Set Mail = obj
        
        'Check if we need to delay delivery
        If SendNow = "N" Then
        
            'Save a temporary sendDate before opening UserForm
            sendDateTemp = sendDate
            
            'Set Userform boxes to sendDate and sendTime
            UserForm1.Label2.Caption = Format(sendDate, "ddd, m/d/yyyy at hh:mm AM/PM")
            UserForm2.Label1.Caption = Format(sendDateTemp, "ddd, m/d/yyyy")
            UserForm2.Label2.Caption = Format(sendDateTemp, "hh:mm AM/PM")
            
            'Initialize UserForm1
            UserForm1.Show vbModal
        
            If UserDeferOption = 1 Then
                Mail.DeferredDeliveryTime = sendDate
                MsgBox ("Your mail will be sent at: " & Format(sendDate, "ddd, m/d/yyyy at hh:mm AM/PM"))
            End If
            
            If UserDeferOption = 2 Then
                'Send Now
                
            Else
                'Otherwise, do not send
                Cancel = True
                
            End If
        End If
    End If
End If

Exit Sub
'ErrorHandler:
' MsgBox "Error!"
Cancel = True

End Sub



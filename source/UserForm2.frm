VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Send at:"
   ClientHeight    =   2610
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim delayDays As Integer
Dim delayHours As Integer
Private Sub CancelButton_Click()
'Reset sendDateTemp
sendDateTemp = sendDate
SpinButton1.Value = 0
SpinButton2.Value = 0

Unload UserForm2

End Sub
Private Sub SaveButton_Click()
'Save over sendDate with sendDateTemp
sendDate = sendDateTemp
UserForm1.Label2.Caption = Format(sendDateTemp, "ddd, m/d/yyyy at hh:mm AM/PM")

Unload UserForm2
End Sub
Private Sub SpinButton1_Change()

Call UpdateVals

End Sub
Private Sub SpinButton2_Change()

Call UpdateVals

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    ' YOUR CODE HERE (Just copy whatever the close button does)
Call CancelButton_Click

    If CloseMode = vbFormControlMenu Then
        Cancel = False
    End If

End Sub

Private Sub UpdateVals()

'Get new SpinButton value
delayDays = SpinButton1.Value
delayHours = SpinButton2.Value

'Generate a new sendDateTemp
sendDateTemp = DateAdd("d", delayDays, sendDate)
sendDateTemp = DateAdd("n", delayHours * 30, sendDateTemp)

'Update labels
UserForm2.Label1.Caption = Format(sendDateTemp, "ddd, m/d/yyyy")
UserForm2.Label2.Caption = Format(sendDateTemp, "hh:mm AM/PM")

End Sub


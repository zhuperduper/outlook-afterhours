VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Delay sending email?"
   ClientHeight    =   2835
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DelaySend_Click()

UserDeferOption = 1

Unload UserForm1
Unload UserForm2

End Sub


Private Sub ModifyDate_Click()

UserForm2.Label1.Caption = Format(sendDateTemp, "ddd, m/d/yyyy")
UserForm2.Label2.Caption = Format(sendDateTemp, "hh:mm AM/PM")
UserForm2.Show

End Sub

Private Sub SendNow_Click()

UserDeferOption = 2

Unload UserForm1
Unload UserForm2

End Sub

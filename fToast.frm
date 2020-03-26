VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fToast
   ClientHeight    =   2124
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3624
   OleObjectBlob   =   "fToast.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fToast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub WhiteButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Make OK Button appear Green when hovered on

    WhiteButton.Visible = False
    
End Sub

Sub Userform_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Reset Userform buttons to Inactive Status

    WhiteButton.Visible = True
    

End Sub
Sub GreenButton_Click()
'PURPOSE: Make OK Button appear Green when hovered on

    Unload Me
    
End Sub


VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fDangerToast 
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "fDangerToast.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fDangerToast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub OKButtonGreen_Click()

Unload Me

End Sub
Sub OKButtonWhite_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Make OK Button appear Green when hovered on

    OKButtonWhite.Visible = False
    
End Sub

Sub Userform_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Reset Userform buttons to Inactive Status

    OKButtonWhite.Visible = True
    

End Sub

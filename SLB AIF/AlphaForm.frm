VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AlphaForm 
   Caption         =   "AlphaForm"
   ClientHeight    =   6030
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6510
   OleObjectBlob   =   "AlphaForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AlphaForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton8_Click()
AlphaForm.Hide
Unload Me
End Sub

Private Sub CommandButton7_Click()


    
    Call submitAlpha
    AlphaForm.Hide
    Unload Me

End Sub

Private Sub TextBox12_Change()

End Sub

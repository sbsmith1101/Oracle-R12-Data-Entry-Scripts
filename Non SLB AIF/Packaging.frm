VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Packaging 
   Caption         =   "Packaging"
   ClientHeight    =   6075
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7380
   OleObjectBlob   =   "Packaging.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Packaging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox1_Click()

End Sub

Private Sub CommandButton7_Click()

   'Dim msgValue As VbMsgBoxResult
    
    'msgValue = MsgBox("Do you want to Save the Item?", vbYesNo + vbInfomation, "Confirmation")
    
    'If msgValue = vbNo Then Exit Sub
    
    Call submitPackaging
    'Call PackagingReset
    Packaging.Hide
    Unload Me

End Sub

Private Sub CommandButton8_Click()
Packaging.Hide
Unload Me
End Sub



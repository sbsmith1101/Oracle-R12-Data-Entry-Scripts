VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MoldForm 
   Caption         =   "Molded Form"
   ClientHeight    =   7140
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4950
   OleObjectBlob   =   "MoldForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MoldForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton3_Click()
   Dim msgValue As VbMsgBoxResult
    
    msgValue = MsgBox("Do you want to Save the Item?", vbYesNo + vbInfomation, "Confirmation")
    
    If msgValue = vbNo Then Exit Sub
    
    Call submitMold
    Call MoldFormReset
    MoldForm.Hide
    CompForm.Show
    Unload Me
End Sub


Private Sub CommandButton4_Click()
MoldForm.Hide
Unload Me


End Sub

Private Sub Label11_Click()

End Sub

Private Sub Label12_Click()

End Sub

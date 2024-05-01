VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BomForm 
   Caption         =   "Bom Form"
   ClientHeight    =   7350
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7095
   OleObjectBlob   =   "BomForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BomForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton1_Click()
   Dim msgValue As VbMsgBoxResult
    
    msgValue = MsgBox("Do you want to Save the Item?", vbYesNo + vbInfomation, "Confirmation")
    
    If msgValue = vbNo Then Exit Sub
    
    Call submitBom
    Call BomFormReset
    BomForm.Hide
    MoldForm.Show
    Unload Me
End Sub

Private Sub CommandButton2_Click()
    UserForm_Initialize
    Unload Me
End Sub


Private Sub Label16_Click()

End Sub

Private Sub UserForm_Initialize()
    Call BomFormReset
    Call MoldFormReset
    Call CompFormReset
    

End Sub



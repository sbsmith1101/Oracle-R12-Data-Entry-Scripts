VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TransferForm 
   Caption         =   "Transfer Form"
   ClientHeight    =   7830
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6915
   OleObjectBlob   =   "TransferForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TransferForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
   Dim msgValue As VbMsgBoxResult
    
    msgValue = MsgBox("Do you want to Save the Item?", vbYesNo + vbInfomation, "Confirmation")
    
    If msgValue = vbNo Then Exit Sub
    
    Call submitBomT
    
    TransferForm.Hide
    If ComboBox1 = "Shoot & Ship" Then Call ResetMoldForm
    If ComboBox1 = "Molded Component" Then Call ResetMoldForm
    If ComboBox1 = "Shoot & Ship" Then MoldForm.Show
    If ComboBox1 = "Molded Component" Then MoldForm.Show
    If ComboBox1 = "Assembly" Then CompForm.Show
    If ComboBox1 = "Sub Assembly" Then CompForm.Show
    
    Unload Me
    
    
End Sub

Private Sub CommandButton2_Click()
    UserForm_Initialize
    Unload Me
End Sub


Private Sub Label16_Click()

End Sub

Private Sub Label20_Click()

End Sub

Private Sub UserForm_Initialize()
    Call ResetBomForm
    'Call ResetMoldForm
    Call ResetCompForm
    

End Sub



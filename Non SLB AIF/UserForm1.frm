VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Sif Form"
   ClientHeight    =   7065
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4755
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton1_Click()
   Dim msgValue As VbMsgBoxResult
    
    msgValue = MsgBox("Do you want to Save the Item?", vbYesNo + vbInfomation, "Confirmation")
    
    If msgValue = vbNo Then Exit Sub
    
    Call submit
    Call SifFormReset
End Sub

Private Sub CommandButton2_Click()
    UserForm1.Hide
End Sub

Private Sub Label4_Click()

End Sub

Private Sub Label5_Click()

End Sub

Private Sub Label7_Click()

End Sub

Private Sub Label8_Click()

End Sub

Private Sub TextBox3_Change()

End Sub

Private Sub UserForm_Initialize()
    Call SifFormReset

End Sub


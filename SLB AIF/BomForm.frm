VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BomForm 
   Caption         =   "Bom Form"
   ClientHeight    =   7950
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


Private Sub ComboBox11_Change()
Dim myTable As ListObject
Dim myArray As Variant
If ComboBox11 = "CNL" Then
    Set myTable = Worksheets("TABLES").ListObjects("PlnCNL")
    myArray = myTable.DataBodyRange
    ComboBox12.List = myArray
End If
If ComboBox11 = "GWH" Then
    Set myTable = Worksheets("TABLES").ListObjects("PlnGWH")
    myArray = myTable.DataBodyRange
    ComboBox12.List = myArray
End If
If ComboBox11 = "LVG" Then
    Set myTable = Worksheets("TABLES").ListObjects("PlnLVG")
    myArray = myTable.DataBodyRange
    ComboBox12.List = myArray
End If
If ComboBox11 = "MEX" Then
    Set myTable = Worksheets("TABLES").ListObjects("PlnMEX")
    myArray = myTable.DataBodyRange
    ComboBox12.List = myArray
End If
If ComboBox11 = "SLB" Then
    Set myTable = Worksheets("TABLES").ListObjects("PlnSLB")
    myArray = myTable.DataBodyRange
    ComboBox12.List = myArray
End If
End Sub

Private Sub CommandButton1_Click()
   Dim msgValue As VbMsgBoxResult
    
    msgValue = MsgBox("Do you want to Save the Item?", vbYesNo + vbInfomation, "Confirmation")
    
    If msgValue = vbNo Then Exit Sub
    
    Call submitBom
    
    BomForm.Hide
    If ComboBox1 = "Shoot & Ship" Then Call MoldFormReset
    If ComboBox1 = "Molded Component" Then Call MoldFormReset
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

Private Sub UserForm_Initialize()
    Call BomFormReset
    'Call MoldFormReset
    Call CompFormReset
    

End Sub



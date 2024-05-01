Attribute VB_Name = "ItemFormCostLaunch"

Sub SifFormReset()
' loads form for add item form on aif sheet
Dim iRow As Long

iRow = [Counta(AIF!B:B)]

    With UserForm1

        .ComboBox1.Clear
        .ComboBox1.AddItem "CNL - 107"
        .ComboBox1.AddItem "GWH - 107"
        .ComboBox1.AddItem "LVG - 105"
        .ComboBox1.AddItem "MEX - 104"
        .ComboBox1.AddItem "SLB-U40 - 109"
        .ComboBox1.AddItem "SLB-U41 - 109"
        .ComboBox1.AddItem "SLB-U42 - 109"
        .ComboBox1.AddItem "SLB-U43 - 109"
        
        .ComboBox2.Clear
        .ComboBox2.AddItem "Pending"
        .ComboBox2.AddItem "Kickoff"
        .ComboBox2.AddItem "Transfer"
        
    
        .ComboBox3.Clear
        .ComboBox3.AddItem "Mold"
        .ComboBox3.AddItem "Assm"
        .ComboBox3.AddItem "Insert-Mld"
        
        .ComboBox4.Clear
        .ComboBox4.AddItem "Transfer"
        .ComboBox4.AddItem "Kickoff"
        .ComboBox4.AddItem "Pending"
        .ComboBox4.AddItem "PassThru"
        .ComboBox4.AddItem "Outsource"
        .ComboBox4.AddItem "CriticalPart"
        .ComboBox4.AddItem "Blend"
        
        
    
    End With

End Sub

Sub submit()
'load info from add item from into AIF sheet item list
    Dim sh As Worksheet
    Dim iRow As Range
    Dim rng As Range
    Set sh = ThisWorkbook.Sheets("AIF")
    Application.FindFormat.Clear
    Set rng = ThisWorkbook.Worksheets("AIF").Range("B5:B40")
    Set iRow = rng.Find(what:="", searchFormat:=False)
    'iRow = [Counta(AIF!B:B)] + 1
    Dim stg As String
    stg = UserForm1.ComboBox1.Value
    
    With sh
    
        .Cells((iRow.Row), 2) = UserForm1.TextBox2.Value
        .Cells((iRow.Row), 3) = Left$(UserForm1.ComboBox1.Value, InStr(stg, " ") - 1)
        .Cells((iRow.Row), 4) = Right$(UserForm1.ComboBox1.Value, 3)
        .Cells((iRow.Row), 5) = UserForm1.ComboBox2.Value
        .Cells((iRow.Row), 6) = UserForm1.ComboBox3.Value
        .Cells((iRow.Row), 7) = UserForm1.ComboBox4.Value
        .Cells((iRow.Row), 8) = UserForm1.TextBox5.Value
        .Cells((iRow.Row), 10) = UserForm1.TextBox3.Value
        .Cells((iRow.Row), 13) = UserForm1.TextBox4.Value
        .Cells((iRow.Row), 11).Interior.ColorIndex = -4142
    End With
        
    
End Sub
Sub submitItemCost()
'load info from add item from into AIF sheet item list
    Dim sh As Worksheet
    Dim iRow As Range
    Dim rng As Range
    Set sh = ThisWorkbook.Sheets("AIF")
    Application.FindFormat.Clear
    Set rng = ThisWorkbook.Worksheets("AIF").Range("B5:B40")
    Set iRow = rng.Find(what:="", searchFormat:=False)
    'iRow = [Counta(AIF!B:B)] + 1
    Dim Itype As String
    Dim DSPS As String
    Dim RGR As String
    Dim RMEX As String
    Dim GLS As String
    Dim Unit As String
    Dim rylt As String
    Dim CostArray As Variant
    
    
    
    If Itype = "Shoot Ship" Then CLShoot
    If Itype = "Mold Comp" Then MComp
    If Itype = "Sub Assy" Then SubA
    If Itype = "Assembly" Then Assm
    
    
    With sh
    
        .Cells((iRow.Row), 18) = Item_Cost_Line
        .Cells((iRow.Row), 3) = Left$(UserForm1.ComboBox1.Value, InStr(stg, " ") - 1)
        .Cells((iRow.Row), 4) = Right$(UserForm1.ComboBox1.Value, 3)
        .Cells((iRow.Row), 5) = UserForm1.ComboBox2.Value
        .Cells((iRow.Row), 6) = UserForm1.ComboBox3.Value
        .Cells((iRow.Row), 7) = UserForm1.ComboBox4.Value
        .Cells((iRow.Row), 8) = UserForm1.TextBox5.Value
        .Cells((iRow.Row), 10) = UserForm1.TextBox3.Value
        .Cells((iRow.Row), 13) = UserForm1.TextBox4.Value
        .Cells((iRow.Row), 11).Interior.ColorIndex = -4142
    End With
    CostArray = Split(.Cells((iRow.Row), 18).Value, Chr(10))
    
End Sub

Sub CLShoot()

.Cells((iRow.Row), 18) = "Yeild" & Unit & "*"
If Unit = "U18" Then .Cells((iRow.Row), 18) = .Cells((iRow.Row), 18) & "0.2"
 Else: .Cells((iRow.Row), 18) = .Cells((iRow.Row), 18) & "0.1"
End If
.Cells((iRow.Row), 18) = .Cells((iRow.Row), 18) & "*" & "Rylty" & Unit & "*" & Item_Cost_Line.TextBox4.Value



End Sub

Sub Show_FormITEMCOST()
'shows userform for add item form, aif sheet
    Item_Cost_Line.Show
End Sub

Sub Show_Form()
'shows userform for add item form, aif sheet
    UserForm1.Show
End Sub

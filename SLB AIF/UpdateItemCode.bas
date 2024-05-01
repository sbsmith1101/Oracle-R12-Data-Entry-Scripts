Attribute VB_Name = "UpdateItemCode"
Public Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As LongPtr, ByVal Y As LongPtr) As LongPtr
Public Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwmilliseconds As LongPtr)

Public Declare PtrSafe Sub mouse_event Lib "user32" (ByVal dwFlags As LongPtr, ByVal dx As Long, ByVal dy As LongPtr, ByVal cButtons As LongPtr, ByVal swextrainfo As LongPtr) '
Public Const mouseeventf_Leftdown = &H2
Public Const mouseeventf_Leftup = &H4
Public Const mouseeventf_Rightdown As Long = &H8
Public Const mouseeventf_rightup As Long = &H10

Public Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Public Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Public Declare PtrSafe Function CloseClipboard Lib "user32" () As Long

Public iRow As Long
Public txt1 As String
Public opRng As Range
Public findCell As Range
Public orgCode As String
Public ItemNum As String
Public toolNum As String
Public partType As String
Public t As String
Public DEPT As String
Public pressRate As String
Public pressSize As String
Public pressCode As String
Public PPH As Integer
Public inspecLvl As Integer
Public inspecCode As String
Public mspakLvl As Integer
Public mspakCode As String
Public gateLvl As Integer
Public gateCode As String
Public AnnealLvl As Integer
Public AnnealCode As String
Public rateCode As String
Public ToolCost As String
Public Planner As String
Public compNum As String
Public SEQ(4) As Variant
Public LABORARR(2) As Variant
Public PACKARR(2) As Variant
Public INSARR(2) As Variant
Public GATEARR(2) As Variant
Public ANNEALARR(2) As Variant


Sub ClearClipboard()
    OpenClipboard (0&)
    EmptyClipboard
    CloseClipboard
End Sub
Sub CopyCompareCell()
'this sub activates a crtl C action and then pulls the info in the windows clipboard and stores it in a variable so it can be evaluated at various intervals
Dim ClipObj As New DataObject

ClipObj.SetText Text:=Empty
ClipObj.PutInClipboard
SendKeys ("^c")
DoEvents
Application.Wait (Now + TimeValue("00:00:01"))


ClipObj.GetFromClipboard
On Error Resume Next
    txt1 = ClipObj.GetText(1)
On Error GoTo 0

End Sub
Sub updateROUT()

Cells(1, 38) = Cells(findCell.Row, 38)
Cells(1, 39) = Cells(findCell.Row, 39)
Cells(1, 40) = Cells(findCell.Row, 40)

For x = 1 To 5
    If InStr(1, Cells(findCell.Row, 39), Chr(10)) = "0" Then GoTo ENext
    'InStr(1, Cells(findCell.Row, 39), Chr(10)) 'position of end of 1st line
    'Left(Cells(findCell.Row, 39), InStr(1, Cells(findCell.Row, 39), Chr(10))-1) first code
    'Cells(findCell.Row, 39) = Mid(Cells(findCell.Row, 39), InStr(1, Cells(findCell.Row, 39), Chr(10)) - 1)
    'MsgBox Mid(Cells(8, 39), InStr(1, Cells(8, 39), Chr(10)) + 1) 'after first line
    'MsgBox Left(Cells(8, 39), InStr(1, Cells(8, 39), Chr(10)) - 1) ' first line
    'MsgBox InStr(1, Cells(9, 39), Chr(10))
    'MsgBox InStr(1, Cells(8, 39), Chr(10))
    
    'cells(findcell.Row,39)
Dim myTable As ListObject
Set myTable = Worksheets("TABLES").ListObjects("RRates")
Dim r As String

r = Application.Match(Left(Cells(findCell.Row, 39), InStr(1, Cells(findCell.Row, 39), Chr(10)) - 1), Worksheets("TABLES").ListObjects("RRates").ListColumns(1).DataBodyRange, 0)


    If Worksheets("TABLES").ListObjects("RRates").DataBodyRange(r, 2) = "Press/Labor" Then
        LABORARR(0) = Left(Cells(findCell.Row, 38), InStr(1, Cells(findCell.Row, 38), Chr(10)) - 1)
        
    End If
    If Worksheets("TABLES").ListObjects("RRates").DataBodyRange(r, 2) = "PACK" Then
        PACKARR(0) = Left(Cells(findCell.Row, 38), InStr(1, Cells(findCell.Row, 38), Chr(10)) - 1)
        
    End If
    If Worksheets("TABLES").ListObjects("RRates").DataBodyRange(r, 2) = "ANNEAL" Then
        ANNEALARR(0) = Left(Cells(findCell.Row, 38), InStr(1, Cells(findCell.Row, 38), Chr(10)) - 1)
        
    End If
    If Worksheets("TABLES").ListObjects("RRates").DataBodyRange(r, 2) = "GATE" Then
        GATEARR(0) = Left(Cells(findCell.Row, 38), InStr(1, Cells(findCell.Row, 38), Chr(10)) - 1)
        
    End If
    If Worksheets("TABLES").ListObjects("RRates").DataBodyRange(r, 2) = "INSPECTION" Then
        INSARR(0) = Left(Cells(findCell.Row, 38), InStr(1, Cells(findCell.Row, 38), Chr(10)) - 1)
        
    End If

    Cells(findCell.Row, 38) = Mid(Cells(findCell.Row, 38), InStr(1, Cells(findCell.Row, 38), Chr(10)) + 1)
    Cells(findCell.Row, 39) = Mid(Cells(findCell.Row, 39), InStr(1, Cells(findCell.Row, 39), Chr(10)) + 1)
    Cells(findCell.Row, 40) = Mid(Cells(findCell.Row, 40), InStr(1, Cells(findCell.Row, 40), Chr(10)) + 1)

Next
Dim vx As Integer
vx = 0
Dim TempSEQ(4) As Variant


ENext:

LABORARR(1) = pressRate
LABORARR(2) = PPH

PACKARR(1) = mspakCode
If mspakLvl = "5" Then
    PACKARR(2) = "5"
Else
    PACKARR(2) = "1"
End If

INSARR(1) = inspecCode
If inspecLvl = "5" Then
    INSARR(2) = "5"
Else
    INSARR(2) = "1"
End If

ANNEALARR(1) = AnnealCode
If AnnealLvl = "5" Then
    ANNEALARR(2) = "5"
Else
    ANNEALARR(2) = "1"
End If

GATEARR(1) = gateCode
If gateLvl = "5" Then
    GATEARR(2) = "5"
Else
    GATEARR(2) = "1"
End If


If Not LABORARR(1) = Empty Then
        LABORARR(0) = Worksheets("TABLES").Cells(2, 2).Offset(vx, 0)
        vx = vx + 1
    If LABORARR(0) = Empty Then
            vx = vx - 1
           LABORARR(0) = Worksheets("TABLES").Cells(2, 2).Offset(vx - 1, 0) + 10
           Worksheets("TABLES").Cells(2, 2).Offset(vx, 0) = LABORARR(0)
           vx = vx + 1
    End If
End If

If Not PACKARR(1) = Empty Then
        PACKARR(0) = Worksheets("TABLES").Cells(2, 2).Offset(vx, 0)
        vx = vx + 1
    If PACKARR(0) = Empty Then
            vx = vx - 1
           PACKARR(0) = Worksheets("TABLES").Cells(2, 2).Offset(vx - 1, 0) + 10
           Worksheets("TABLES").Cells(2, 2).Offset(vx, 0) = PACKARR(0)
           vx = vx + 1
    End If
End If

If Not ANNEALARR(1) = Empty Then
        ANNEALARR(0) = Worksheets("TABLES").Cells(2, 2).Offset(vx, 0)
        vx = vx + 1
    If ANNEALARR(0) = Empty Then
            vx = vx - 1
           ANNEALARR(0) = Worksheets("TABLES").Cells(2, 2).Offset(vx - 1, 0) + 10
           Worksheets("TABLES").Cells(2, 2).Offset(vx, 0) = ANNEALARR(0)
           vx = vx + 1
    End If
End If

If Not GATEARR(1) = Empty Then
        GATEARR(0) = Worksheets("TABLES").Cells(2, 2).Offset(vx, 0)
        vx = vx + 1
   
    If GATEARR(0) = Empty Then
            vx = vx - 1
           GATEARR(0) = Worksheets("TABLES").Cells(2, 2).Offset(vx - 1, 0) + 10
           Worksheets("TABLES").Cells(2, 2).Offset(vx, 0) = GATEARR(0)
           vx = vx + 1
    End If
End If

If Not INSARR(1) = Empty Then
        INSARR(0) = Worksheets("TABLES").Cells(2, 2).Offset(vx, 0)
        vx = vx + 1
    If INSARR(0) = Empty Then
            vx = vx - 1
           INSARR(0) = Worksheets("TABLES").Cells(2, 2).Offset(vx - 1, 0) + 10
           Worksheets("TABLES").Cells(2, 2).Offset(vx, 0) = INSARR(0)
           vx = vx + 1
    End If
End If
For x = 1 To 5 - vx
  Worksheets("TABLES").Cells(2, 2).Offset(5 - x, 0) = ""
Next


Cells(findCell.Row, 38) = Cells(1, 38)
Cells(findCell.Row, 39) = Cells(1, 39)
Cells(findCell.Row, 40) = Cells(1, 40)


Cells(1, 38) = Empty
Cells(1, 39) = Empty
Cells(1, 40) = Empty
Cells(2, 2) = Empty
Cells(3, 2) = Empty
Cells(4, 2) = Empty
Cells(5, 2) = Empty
Cells(6, 2) = Empty

'find max value of seq

End Sub
Sub enterNEWROUT()


SendKeys ("%vdf")

slow1
Dim count As Integer
count = 0

If Not LABORARR(1) = Empty Then
    count = count + 1
    SendKeys LABORARR(0)
    slow1
    SendKeys ("{Tab}")
    SendKeys LABORARR(1)
    slow1
    SendKeys ("{Tab}")
    SendKeys ("{Tab}")
    SendKeys ("{Tab}")
    SendKeys ("{Tab}")
    slow1
    SendKeys LABORARR(2)
    'SendKeys ("+{pgup}")
    'SendKeys ("+{pgdn}")
    'For x = 1 To count
    'count = count + 1
    '    SendKeys ("{Down}")
    'Next

        SendKeys ("+{Tab}")
        SendKeys ("+{Tab}")
        SendKeys ("+{Tab}")
        SendKeys ("+{Tab}")
        SendKeys ("+{Tab}")
        slow1

    For x = 1 To 3
        CopyCompareCell
        If txt1 = LABORARR(0) Then
            x = x + 1
        Else
            SendKeys ("+{Tab}")
            slow1
        End If
    Next
    SendKeys ("{Down}")
End If


slow1

If Not PACKARR(1) = Empty Then
    count = count + 1
    SendKeys PACKARR(0)
    slow1
    SendKeys ("{Tab}")
    SendKeys PACKARR(1)
     slow1
    SendKeys ("{Tab}")
    SendKeys ("{Tab}")
    SendKeys ("{Tab}")
    SendKeys ("{Tab}")
    slow1
    SendKeys PACKARR(2)
    'SendKeys ("%vi")
    'SendKeys ("{Down}")

        SendKeys ("+{Tab}")
        SendKeys ("+{Tab}")
        SendKeys ("+{Tab}")
        SendKeys ("+{Tab}")
        SendKeys ("+{Tab}")
        slow1

    For x = 1 To 3
        CopyCompareCell
        If txt1 = PACKARR(0) Then
            x = x + 1
        Else
            SendKeys ("+{Tab}")
            slow1
        End If
    Next
    SendKeys ("{Down}")
End If


slow1

If Not ANNEALARR(1) = Empty Then
    count = count + 1
    SendKeys ANNEALARR(0)
    slow1
    SendKeys ("{Tab}")
    SendKeys ANNEALARR(1)
     slow1
    SendKeys ("{Tab}")
    SendKeys ("{Tab}")
    SendKeys ("{Tab}")
    SendKeys ("{Tab}")
    slow1
    SendKeys ANNEALARR(2)
    'SendKeys ("%vi")
    'SendKeys ("{Down}")

        SendKeys ("+{Tab}")
        SendKeys ("+{Tab}")
        SendKeys ("+{Tab}")
        SendKeys ("+{Tab}")
        SendKeys ("+{Tab}")
        slow1

    For x = 1 To 3
        CopyCompareCell
        If txt1 = ANNEALARR(0) Then
            x = x + 1
        Else
            SendKeys ("+{Tab}")
            slow1
        End If
    Next
    SendKeys ("{Down}")
End If


slow1

If Not GATEARR(1) = Empty Then
    count = count + 1
    SendKeys GATEARR(0)
    slow1
    SendKeys ("{Tab}")
    SendKeys GATEARR(1)
     slow1
    SendKeys ("{Tab}")
    SendKeys ("{Tab}")
    SendKeys ("{Tab}")
    SendKeys ("{Tab}")
    slow1
    SendKeys GATEARR(2)
    'SendKeys ("%vi")
    'SendKeys ("{Down}")

        SendKeys ("+{Tab}")
        SendKeys ("+{Tab}")
        SendKeys ("+{Tab}")
        SendKeys ("+{Tab}")
        SendKeys ("+{Tab}")
        slow1

    For x = 1 To 3
        CopyCompareCell
        If txt1 = GATEARR(0) Then
            x = x + 1
        Else
            SendKeys ("+{Tab}")
            slow1
        End If
    Next
    SendKeys ("{Down}")
End If


slow1

If Not INSARR(1) = Empty Then
    count = count + 1
    SendKeys INSARR(0)
    slow1
    SendKeys ("{Tab}")
    SendKeys INSARR(1)
     slow1
    SendKeys ("{Tab}")
    SendKeys ("{Tab}")
    SendKeys ("{Tab}")
    SendKeys ("{Tab}")
    slow1
    SendKeys INSARR(2)
    'SendKeys ("%vi")
    'SendKeys ("{Down}")

        SendKeys ("+{Tab}")
        SendKeys ("+{Tab}")
        SendKeys ("+{Tab}")
        SendKeys ("+{Tab}")
        SendKeys ("+{Tab}")
        slow1

    For x = 1 To 3
        CopyCompareCell
        If txt1 = INSARR(0) Then
            x = x + 1
        Else
            SendKeys ("+{Tab}")
            slow1
        End If
    Next
    SendKeys ("{Down}")
End If
slow1

'remove extra items

count = 5 - count
For x = 1 To count
   SendKeys ("%ed")
    SendKeys ("%o")
    SendKeys ("%o")
    slow1
Next

SendKeys ("%fs")

slow2


'endENT
End Sub

Sub setvars()
'this sub sets assigns the values of item variables which are used by other subs

t = findCell.Column

ItemNum = Cells(findCell.Row, 2).Value
orgCode = Cells(findCell.Row, 22).Value
iRow = findCell.Row
PPH = Cells(findCell.Row, 9).Value
inspecLvl = Cells(findCell.Row, 7).Value
mspakLvl = Cells(findCell.Row, 8).Value
gateLvl = Cells(findCell.Row, 31).Value
AnnealLvl = Cells(findCell.Row, 32).Value
DEPT = Cells(findCell.Row, 23).Value


pressCode = ""
rateCode = ""

pressSize = Cells(findCell.Row, 16).Value

If Cells(findCell.Row, 10).Value = "Automatic" Then pressCode = "AL"
If Cells(findCell.Row, 10).Value = "Semi-Automatic" Then pressCode = "SL"
If Cells(findCell.Row, 19).Value = "Yes" Then pressCode = pressCode & "F"


pressRate = pressCode & Mid(DEPT, 2) & pressSize & "T"
rateCode = DEPT & "-"


End Sub
Sub submitPackagingT()
'take info from packaging form and load it into kickoff sheet
ThisWorkbook.Worksheets("LocArray").Unprotect Password:="1234"
 Dim ph As Worksheet
    Dim text1 As String
    Dim text2 As String
    Dim Counter As Integer
    Dim List As Variant
    Dim Item As Variant
    Dim rng As Range
    Dim LocRng As Range
    Set LocRng = ThisWorkbook.Worksheets("LocArray").Range("x3:x7")
    
    
    Set ph = ThisWorkbook.Sheets("Transfer Line Calculator")

    ThisWorkbook.Sheets("LocArray").Cells(3, 24) = Packaging.TextBox1.Value
    ThisWorkbook.Sheets("LocArray").Cells(4, 24) = Packaging.TextBox2.Value
    ThisWorkbook.Sheets("LocArray").Cells(5, 24) = Packaging.TextBox3.Value
    ThisWorkbook.Sheets("LocArray").Cells(6, 24) = Packaging.TextBox4.Value
    ThisWorkbook.Sheets("LocArray").Cells(7, 24) = Packaging.TextBox5.Value
    
    
    ThisWorkbook.Sheets("LocArray").Cells(3, 25) = Packaging.TextBox6.Value
    ThisWorkbook.Sheets("LocArray").Cells(4, 25) = Packaging.TextBox7.Value
    ThisWorkbook.Sheets("LocArray").Cells(5, 25) = Packaging.TextBox8.Value
    ThisWorkbook.Sheets("LocArray").Cells(6, 25) = Packaging.TextBox9.Value
    ThisWorkbook.Sheets("LocArray").Cells(7, 25) = Packaging.TextBox10.Value
' this later section is probably obsolete, check later
    
    For Each rng In LocRng  ' take list and create lines in display sheet
        If rng = "" Then GoTo Skip
        
    text1 = text1 & rng.Value & vbCrLf
    text2 = text2 & rng.Offset(0, 1).Value & vbCrLf

Skip:
    Next rng
    If text1 = "" Then GoTo EndPoint
    text1 = Left(text1, Len(text1) - 2)
    text2 = Left(text2, Len(text2) - 2)
    
    
    Cells(iRow, 24) = text1
    Cells(iRow, 25) = text2
    
    
    ' Post information to line
    'With ph
        'Cells(iRow, 26) = CompForm.TextBox1.Value & vbCrLf & CompForm.TextBox2.Value & vbCrLf & CompForm.TextBox3.Value & vbCrLf & CompForm.TextBox4.Value & vbCrLf & CompForm.TextBox5.Value & vbCrLf & _
        'CompForm.TextBox6.Value & vbCrLf & CompForm.TextBox7.Value & vbCrLf & CompForm.TextBox8.Value & vbCrLf & CompForm.TextBox9.Value & vbCrLf & CompForm.TextBox10.Value
        
        
        
    'End With
EndPoint:
ThisWorkbook.Worksheets("LocArray").Protect Password:="1234"
End Sub
Sub submitBomT()
'this takes info from userform and loads it into Kickoff sheet

Set opRng = Worksheets("Transfer Line Calculator").Range("B9:B35")

iRow = Cells(Rows.count, 2).End(xlUp).Row + 1

    Dim ph As Worksheet
    
    Set ph = ThisWorkbook.Sheets("Transfer Line Calculator")
 
    
    With ph
        
        
        
        .Cells(iRow, 1) = Cells(iRow, 1).Offset(-1, 0) + 1
        .Cells(iRow, 2) = TransferForm.TextBox2.Value
        .Cells(iRow, 3) = TransferForm.TextBox3.Value
        .Cells(iRow, 4) = TransferForm.ComboBox1.Value
        .Cells(iRow, 5) = TransferForm.TextBox4.Value
        .Cells(iRow, 6) = TransferForm.ComboBox10.Value
        .Cells(iRow, 7) = TransferForm.ComboBox2.Value
        .Cells(iRow, 8) = TransferForm.ComboBox3.Value
        .Cells(iRow, 9) = TransferForm.TextBox5.Value
        .Cells(iRow, 10) = TransferForm.ComboBox4.Value
        .Cells(iRow, 11) = TransferForm.TextBox6.Value
        .Cells(iRow, 12) = TransferForm.TextBox7.Value
        .Cells(iRow, 13) = TransferForm.TextBox6.Value
        .Cells(iRow, 15) = TransferForm.TextBox8.Value
        .Cells(iRow, 22) = TransferForm.ComboBox11.Value
        .Cells(iRow, 23) = TransferForm.ComboBox14.Value
        .Cells(iRow, 34) = TransferForm.TextBox9.Value
        
    End With
        
        
End Sub
Sub submitMoldT()
'take info from mold form and loads it into Kickoff sheet
    Dim xRng As Range, xCell As Range
    Dim i As Integer
    Dim th As Worksheet
    Dim text1 As String
    Dim text2 As String
    
    Set th = ThisWorkbook.Sheets("Transfer Line Calculator")

    
    With th
        
        If MoldForm.CheckBox1.Value = True Then Cells(iRow, 18) = "Yes"
        If MoldForm.CheckBox2.Value = True Then Cells(iRow, 19) = "Yes"
        
        
        text1 = MoldForm.TextBox1.Value & vbCrLf & MoldForm.TextBox3.Value
        If MoldForm.TextBox3.Value = "" Then text1 = Left(text1, Len(text1) - 2)
        Cells(iRow, 20) = text1

        text2 = MoldForm.TextBox2.Value & vbCrLf & MoldForm.TextBox4.Value
        If MoldForm.TextBox4.Value = "" Then text2 = Left(text2, Len(text2) - 2)
        Cells(iRow, 21) = text2
        
        .Cells(iRow, 13) = MoldForm.TextBox5.Value
        .Cells(iRow, 14) = MoldForm.TextBox6.Value
        .Cells(iRow, 16).NumberFormat = "@"
        .Cells(iRow, 16) = MoldForm.ComboBox2.Value
        .Cells(iRow, 17) = MoldForm.ComboBox1.Value
        .Cells(iRow, 31) = MoldForm.ComboBox3.Value
        .Cells(iRow, 32) = MoldForm.ComboBox4.Value
        .Cells(iRow, 33) = MoldForm.TextBox7.Value

    End With
        
        
End Sub
Sub submitCompT()
' this takes info from comp form and load it into kickoff sheet
ThisWorkbook.Worksheets("LocArray").Unprotect Password:="1234"
 Dim ph As Worksheet
    Dim text1 As String
    Dim text2 As String
    Dim text3 As String
    Dim text4 As String
    Dim rng As Range
    Dim LocRng As Range
    Set LocRng = ThisWorkbook.Worksheets("LocArray").Range("s3:s12")
    
    
    Set ph = ThisWorkbook.Sheets("Transfer Line Calculator")
    If CompForm.CheckBox1.Value = True Then Cells(iRow, 35) = "Yes"

    ThisWorkbook.Sheets("LocArray").Cells(3, 19) = CompForm.TextBox1.Value
    ThisWorkbook.Sheets("LocArray").Cells(4, 19) = CompForm.TextBox2.Value
    ThisWorkbook.Sheets("LocArray").Cells(5, 19) = CompForm.TextBox3.Value
    ThisWorkbook.Sheets("LocArray").Cells(6, 19) = CompForm.TextBox4.Value
    ThisWorkbook.Sheets("LocArray").Cells(7, 19) = CompForm.TextBox5.Value
    ThisWorkbook.Sheets("LocArray").Cells(8, 19) = CompForm.TextBox6.Value
    ThisWorkbook.Sheets("LocArray").Cells(9, 19) = CompForm.TextBox7.Value
    ThisWorkbook.Sheets("LocArray").Cells(10, 19) = CompForm.TextBox8.Value
    ThisWorkbook.Sheets("LocArray").Cells(11, 19) = CompForm.TextBox9.Value
    ThisWorkbook.Sheets("LocArray").Cells(12, 19) = CompForm.TextBox10.Value
    
    
    ThisWorkbook.Sheets("LocArray").Cells(3, 20) = CompForm.TextBox31.Value
    ThisWorkbook.Sheets("LocArray").Cells(4, 20) = CompForm.TextBox32.Value
    ThisWorkbook.Sheets("LocArray").Cells(5, 20) = CompForm.TextBox33.Value
    ThisWorkbook.Sheets("LocArray").Cells(6, 20) = CompForm.TextBox34.Value
    ThisWorkbook.Sheets("LocArray").Cells(7, 20) = CompForm.TextBox35.Value
    ThisWorkbook.Sheets("LocArray").Cells(8, 20) = CompForm.TextBox36.Value
    ThisWorkbook.Sheets("LocArray").Cells(9, 20) = CompForm.TextBox37.Value
    ThisWorkbook.Sheets("LocArray").Cells(10, 20) = CompForm.TextBox38.Value
    ThisWorkbook.Sheets("LocArray").Cells(11, 20) = CompForm.TextBox39.Value
    ThisWorkbook.Sheets("LocArray").Cells(12, 20) = CompForm.TextBox40.Value
    
    
    ThisWorkbook.Sheets("LocArray").Cells(3, 21) = CompForm.TextBox41.Value
    ThisWorkbook.Sheets("LocArray").Cells(4, 21) = CompForm.TextBox42.Value
    ThisWorkbook.Sheets("LocArray").Cells(5, 21) = CompForm.TextBox43.Value
    ThisWorkbook.Sheets("LocArray").Cells(6, 21) = CompForm.TextBox44.Value
    ThisWorkbook.Sheets("LocArray").Cells(7, 21) = CompForm.TextBox45.Value
    ThisWorkbook.Sheets("LocArray").Cells(8, 21) = CompForm.TextBox46.Value
    ThisWorkbook.Sheets("LocArray").Cells(9, 21) = CompForm.TextBox47.Value
    ThisWorkbook.Sheets("LocArray").Cells(10, 21) = CompForm.TextBox48.Value
    ThisWorkbook.Sheets("LocArray").Cells(11, 21) = CompForm.TextBox49.Value
    ThisWorkbook.Sheets("LocArray").Cells(12, 21) = CompForm.TextBox50.Value
    
    
    ThisWorkbook.Sheets("LocArray").Cells(3, 22) = CompForm.ComboBox1.Value
    ThisWorkbook.Sheets("LocArray").Cells(4, 22) = CompForm.ComboBox2.Value
    ThisWorkbook.Sheets("LocArray").Cells(5, 22) = CompForm.ComboBox3.Value
    ThisWorkbook.Sheets("LocArray").Cells(6, 22) = CompForm.ComboBox4.Value
    ThisWorkbook.Sheets("LocArray").Cells(7, 22) = CompForm.ComboBox5.Value
    ThisWorkbook.Sheets("LocArray").Cells(8, 22) = CompForm.ComboBox6.Value
    ThisWorkbook.Sheets("LocArray").Cells(9, 22) = CompForm.ComboBox7.Value
    ThisWorkbook.Sheets("LocArray").Cells(10, 22) = CompForm.ComboBox8.Value
    ThisWorkbook.Sheets("LocArray").Cells(11, 22) = CompForm.ComboBox9.Value
    ThisWorkbook.Sheets("LocArray").Cells(12, 22) = CompForm.ComboBox10.Value

    
    For Each rng In LocRng  ' take list and create lines in display sheet
        If rng = "" Then GoTo Skip
        
        
    
    'If empty skip to next
    text1 = text1 & rng.Value & vbCrLf
    text2 = text2 & rng.Offset(0, 1).Value & vbCrLf
    text3 = text3 & rng.Offset(0, 2).Value & vbCrLf
    text4 = text4 & rng.Offset(0, 3).Value & vbCrLf

Skip:
    Next rng
    
    If text1 = "" Then GoTo EndPoint
    text1 = Left(text1, Len(text1) - 2)
    text2 = Left(text2, Len(text2) - 2)
    text3 = Left(text3, Len(text3) - 2)
    text4 = Left(text4, Len(text4) - 2)
    
    Cells(iRow, 26) = text1
    Cells(iRow, 27) = text2
    Cells(iRow, 28) = text3
    Cells(iRow, 29) = text4
    
    ' Post information to line
    'With ph
        'Cells(iRow, 26) = CompForm.TextBox1.Value & vbCrLf & CompForm.TextBox2.Value & vbCrLf & CompForm.TextBox3.Value & vbCrLf & CompForm.TextBox4.Value & vbCrLf & CompForm.TextBox5.Value & vbCrLf & _
        'CompForm.TextBox6.Value & vbCrLf & CompForm.TextBox7.Value & vbCrLf & CompForm.TextBox8.Value & vbCrLf & CompForm.TextBox9.Value & vbCrLf & CompForm.TextBox10.Value
        
        
        
    'End With
EndPoint:
ThisWorkbook.Worksheets("LocArray").Protect Password:="1234"
End Sub

Sub ShowBom_FormT()
'call for BOM form
    TransferForm.Show
End Sub
Sub ShowMold_FormT()
'call for Mold form
    MoldForm.Show
End Sub
Sub ShowComp_FormT()
'call for Comp form
    CompForm.Show
End Sub
Sub GrabToolCost()

'this sub grabs the acct code from the LocArray sheet, the accounting code based on item type and org

Dim RNG1 As Range
Dim RNG2 As Range
Dim findCellRow As Range
Dim OrgCol As Range
Set RNG1 = ThisWorkbook.Worksheets("LocArray").Range("i17:i21")
Set RNG2 = ThisWorkbook.Worksheets("LocArray").Range("j16:n16")

Set findCellRow = RNG1.Find(what:=Cells(findCell.Row, 17).Value, MatchCase:=False)

Set OrgCol = RNG2.Find(what:=orgCode, MatchCase:=False)

DoEvents

If OrgCol Is Nothing Then ExitCon = True
If OrgCol Is Nothing Then Exit Sub

If Not Cells(findCell.Row, 17).Value = "" Then
If Cells(findCell.Row, 36).Value = "MEX" Then
 ToolCost = Worksheets("LocArray").Cells(findCellRow.Row, 13).Value
Else
ToolCost = Worksheets("LocArray").Cells(findCellRow.Row, OrgCol.Column).Value
End If
End If
DoEvents
End Sub

Sub LogResinsT()

 Dim sh As Worksheet
    Dim iRow As Range
    Dim rng As Range
    Set sh = ThisWorkbook.Sheets("CompCostUpdate")
    Application.FindFormat.Clear
    Set rng = ThisWorkbook.Worksheets("CompCostUpdate").Range("A2:A80")
    Set iRow = rng.Find(what:="", searchFormat:=False)
    'iRow = [Counta(AIF!B:B)] + 1
    
    str1 = ((Replace(Cells(findCell.Row, 20), Chr(10), "/"))) & "/"
    If str1 = "/" Then GoTo SKO
    For x = 1 To 99
        Set iRow = rng.Find(what:="", searchFormat:=False)
        If Not str1 = "" Or str1 = "/" Then ThisWorkbook.Worksheets("CompCostUpdate").Cells((iRow.Row), 1) = Mid(str1, 1, InStr(1, str1, "/") - 1)
        If Not str1 = "" Or str1 = "/" Then ThisWorkbook.Worksheets("CompCostUpdate").Cells((iRow.Row), 2) = Cells(findCell.Row, 22)
        If Not str1 = "" Or str1 = "/" Then ThisWorkbook.Worksheets("CompCostUpdate").Cells((iRow.Row), 6) = Cells(findCell.Row, 36)
        'Stop
        str1 = Mid(str1, InStr(1, str1, "/") + 1)
        
        If str1 = "" Then x = "100"
        
        
    Next
    
SKO:
 
    
End Sub
Sub LogCompsT()

 Dim sh As Worksheet
 Dim str1 As String
    Dim iRow As Range
    Dim rng As Range
    Set sh = ThisWorkbook.Sheets("CompCostUpdate")
    Application.FindFormat.Clear
    Set rng = ThisWorkbook.Worksheets("CompCostUpdate").Range("A2:A80")
    Set iRow = rng.Find(what:="", searchFormat:=False)
    'iRow = [Counta(AIF!B:B)] + 1
    
    str1 = ((Replace(Cells(findCell.Row, 26), Chr(10), "/"))) & "/"
    If str1 = "/" Then GoTo SKO
    For x = 1 To 99
        Set iRow = rng.Find(what:="", searchFormat:=False)
        If Not str1 = "" Or str1 = "/" Then ThisWorkbook.Worksheets("CompCostUpdate").Cells((iRow.Row), 1) = Mid(str1, 1, InStr(1, str1, "/") - 1)
        If Not str1 = "" Or str1 = "/" Then ThisWorkbook.Worksheets("CompCostUpdate").Cells((iRow.Row), 2) = Cells(findCell.Row, 22)
        If Not str1 = "" Or str1 = "/" Then ThisWorkbook.Worksheets("CompCostUpdate").Cells((iRow.Row), 6) = Cells(findCell.Row, 36)
        'Stop
        str1 = Mid(str1, InStr(1, str1, "/") + 1)
        
        If str1 = "" Then x = "100"
        
        
    Next
    
SKO:

    
End Sub
Sub LogPacksT()

 Dim sh As Worksheet
 Dim str1 As String
    Dim iRow As Range
    Dim rng As Range
    Set sh = ThisWorkbook.Sheets("CompCostUpdate")
    Application.FindFormat.Clear
    Set rng = ThisWorkbook.Worksheets("CompCostUpdate").Range("A2:A80")
    Set iRow = rng.Find(what:="", searchFormat:=False)
    'iRow = [Counta(AIF!B:B)] + 1
    
    str1 = ((Replace(Cells(findCell.Row, 24), Chr(10), "/"))) & "/"
    If str1 = "/" Then GoTo SKP
    For x = 1 To 99
        Set iRow = rng.Find(what:="", searchFormat:=False)
        If Not str1 = "" Or str1 = "/" Then ThisWorkbook.Worksheets("CompCostUpdate").Cells((iRow.Row), 1) = Mid(str1, 1, InStr(1, str1, "/") - 1)
        If Not str1 = "" Or str1 = "/" Then ThisWorkbook.Worksheets("CompCostUpdate").Cells((iRow.Row), 2) = Cells(findCell.Row, 22)
        If Not str1 = "" Or str1 = "/" Then ThisWorkbook.Worksheets("CompCostUpdate").Cells((iRow.Row), 6) = Cells(findCell.Row, 36)
        'Stop
        'Stop
        str1 = Mid(str1, InStr(1, str1, "/") + 1)
        
        If str1 = "" Then x = "100"
    Next
    
SKP:
   
End Sub

Sub MainSeqT()


Dim Cl As Range
Dim wrkRng As Range
Set wrkRng = ThisWorkbook.Worksheets("Transfer Line Calculator").Range("B8:B113")



Application.FindFormat.Clear

Application.FindFormat.Interior.ColorIndex = 2

Dim rng As Range
Set rng = ThisWorkbook.Worksheets("Transfer Line Calculator").Range("B8:B113")
Set findCell = rng.Find(what:="*", searchFormat:=True)
If (findCell Is Nothing) Then
    slow1
    DoEvents
    BringToFrontT
    MsgBox "No Unprocessed Items"
    End
    Exit Sub
    

End If

'If Not Cells(findCell.Row, 39).Value = "" Then GoTo ReStart


LogCompsT
LogPacksT
LogResinsT
removedupComps

Stop
entercostT

RoutingLoad
RunBOM

TransferFinishtoAIFSheetT
findCell.Interior.ColorIndex = 4
slow1
GoTo EndLine
ReStart:
ReStartSeq

EndLine:
DoEvents
End Sub
Sub ReStartSeq()

If Cells(findCell.Row, 40).Value = "entercost" Then
    
    entercost
End If
If Cells(findCell.Row, 40).Value = "RunBOM" Then
    
    RunBOM
    entercost
End If
On Error Resume Next
If Cells(findCell.Row, 40).Value = "RoutingLoad" Then
    RoutingLoad
    RunBOM
    entercost
End If
On Error GoTo 0
slow1
DoEvents
End Sub
Sub TransferFinishtoAIFSheetT()

 Dim sh As Worksheet
    Dim iRow As Range
    Dim rng As Range
    Set sh = ThisWorkbook.Sheets("AIF")
    Application.FindFormat.Clear
    Set rng = ThisWorkbook.Worksheets("AIF").Range("B5:B40")
    Set iRow = rng.Find(what:="", searchFormat:=False)
    'iRow = [Counta(AIF!B:B)] + 1
    
    
    With sh
    
        .Cells((iRow.Row), 2) = Cells(findCell.Row, 2).Value
        If Cells(findCell.Row, 4) = "Assembly" Then .Cells((iRow.Row), 6) = "Assm"
        If Cells(findCell.Row, 4) = "Shoot & Ship" Then .Cells((iRow.Row), 6) = "Mold"
        If Cells(findCell.Row, 4) = "Molded Component" Then .Cells((iRow.Row), 6) = "Mold"
        .Cells((iRow.Row), 5) = Cells(findCell.Row, 36)
        .Cells((iRow.Row), 7) = Cells(findCell.Row, 36)
        If Cells(findCell.Row, 22) = "CNL" Then
            .Cells((iRow.Row), 3) = "CNL"
            .Cells((iRow.Row), 4) = "107"
        End If
        If Cells(findCell.Row, 22) = "GWH" Then
            .Cells((iRow.Row), 3) = "GWH"
            .Cells((iRow.Row), 4) = "107"
        End If
        If Cells(findCell.Row, 22) = "LVG" Then
            .Cells((iRow.Row), 3) = "LVG"
            .Cells((iRow.Row), 4) = "105"
        End If
        If Cells(findCell.Row, 22) = "MEX" Then
            .Cells((iRow.Row), 3) = "MEX"
            .Cells((iRow.Row), 4) = "104"
        End If
        If Cells(findCell.Row, 22) = "SLB" Then
            .Cells((iRow.Row), 3) = "SLB"
            .Cells((iRow.Row), 4) = "109"
        End If
        .Cells((iRow.Row), 13) = Cells(findCell.Row, 12).Value
    End With
        


End Sub


Sub codeGrabber()
'this sub grabs rates from tables on the Locarray sheet, based on org and lvl

'pressRate = Worksheets("LocArray").Cells(PressRow.Row, pressCol.Column).Value
If Not inspecLvl = Empty Then
    inspecCode = "Insp" & Mid(DEPT, 2) & "-" & inspecLvl
End If
If Not mspakLvl = Empty Then
    mspakCode = "Ms&Pk" & Mid(DEPT, 2) & "-" & mspakLvl
End If
If Not gateLvl = Empty Then
    gateCode = "GtCt" & Mid(DEPT, 2) & "-" & gateLvl
End If
If Not AnnealLvl = Empty Then
    AnnealCode = "Annl" & Mid(DEPT, 2) & "-" & AnnealLvl
End If

End Sub
Sub changeOrg()
' this sub handles change of orgs in oracle
Dim c As String
If orgCode = "CNL" Then c = "cn"
If orgCode = "GWH" Then c = "g"
If orgCode = "LVG" Then c = "l"
If orgCode = "MEX" Then c = "Me"
If orgCode = "SLB" Then c = "s"
slow1

SendKeys ("%tu"), True
slow1
SendKeys ("o"), True
SendKeys ("%o"), True
SendKeys ("c"), True
slow1
SendKeys ("+{Tab}")
CopyCompareCell
If Not txt1 = "%" Then MsgBox "Out of place error, stoping script"
If Not txt1 = "%" Then End
SendKeys ("{Tab}")

SendKeys (c), True
slow1

End Sub
Public Sub ClickOnCornerWindow()
'sub for clicking on oracle app window to bring it into focus, if I can figure out setfocus method I might be able to eliminate this, currently faces window resizing issues
Dim oLeft As Long
'oLeft = 100  'last desktop size
oLeft = 80   'revised desktop

Dim OTop As Long
OTop = 100

SetCursorPos oLeft, OTop

mouse_event mouseeventf_Leftdown, 0, 0, 0, 0
mouse_event mouseeventf_Leftup, 0, 0, 0, 0


End Sub
Sub slow1()
'sub for adding delay in procedures, delays needed to keep step from overrunning oracle and causing misalignment.
DoEvents
Sleep 1000

End Sub
Sub slow2()
'2 sec version of former sub
DoEvents
Sleep 2000

End Sub
Sub changeBM()
'this is for changing orgs in Bom and routing section
orgCode = Cells(findCell.Row, 22).Value
slow1
SendKeys ("%to"), True
SendKeys ("o"), True
SendKeys ("c"), True
slow1
SendKeys ("%o"), True
slow1

Dim c As String
If orgCode = "CNL" Then c = "cn"
If orgCode = "GWH" Then c = "g"
If orgCode = "LVG" Then c = "l"
If orgCode = "MEX" Then c = "Me"
If orgCode = "SLB" Then c = "s"
slow1
SendKeys ("+{Tab}")
CopyCompareCell
If Not txt1 = "%" Then MsgBox "Out of place error, stoping script"
If Not txt1 = "%" Then End
SendKeys ("{Tab}")

SendKeys (c), True
slow1


End Sub
Sub BringToFrontT()
'brings kickoff boms sheet into focus
    Dim setFocus As Long
    
    ThisWorkbook.Worksheets("Transfer Line Calculator").Activate
    setFocus = SetForegroundWindow(Application.hWnd)
End Sub
Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
  Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i) = stringToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False

End Function
Sub RoutMap()

ClickOnCornerWindow


ThisWorkbook.Worksheets("Transfer Line Calculator").Activate
Application.FindFormat.Clear
Application.FindFormat.Interior.ColorIndex = 2

Dim rng As Range
Set rng = ThisWorkbook.Worksheets("Transfer Line Calculator").Range("B7:B113")
Set findCell = rng.Find(what:="*", searchFormat:=True)

slow1
ClickOnCornerWindow
setvars
codeGrabber
slow2
SendKeys ("%fw"), True

slow1
SendKeys ("nb"), True
slow1


changeBM
SendKeys ("%tc"), True

slow1
SendKeys ("%tu"), True
slow1
SendKeys ("r"), True
SendKeys ("~")
SendKeys ("~")
'enter item and check 00 to ensure we don't encouter item alread exsist error
slow1
SendKeys ("%vf"), True
slow1
SendKeys ItemNum
SendKeys ("%i"), True
slow1
SendKeys "{tab}"
SendKeys "{tab}"
SendKeys "{tab}"

CopyCompareCell

If Not txt1 = "00" Then BringToFrontT
If Not txt1 = "00" Then MsgBox "out of alignment"
If Not txt1 = "00" Then End
DoEvents

slow1
SendKeys "+{PGDN}"
slow1
SendKeys ("%vf"), True
SendKeys ("{Tab}")
slow1
SendKeys ("RATES"), True
SendKeys ("%i"), True
slow1
SendKeys ("%r"), True


Dim SEQR As Range
Set SEQR = Range(Worksheets("TABLES").Cells(2, 2), Worksheets("TABLES").Cells(2, 6))
Dim itemseq() As Variant
Dim u As Integer
u = 0
Dim RArray() As Variant
Dim IArray() As Variant
slow1
CopyCompareCell
    If txt1 = "" Then BringToFrontT
    If txt1 = "" Then MsgBox "Rout Map out of alignment"
    If txt1 = "" Then Stop
    
For Z = 0 To 4
    CopyCompareCell
    If txt1 = "" Then ReDim Preserve itemseq(0 To Z - 1)
    If txt1 = "" Then GoTo SKP1
    If Not txt1 = "" Then
    u = u + 1
    ReDim Preserve itemseq(0 To Z) ' Redimension:
    'itemseq(UBound(itemseq)) = txt1 ' Fill last element
    itemseq(Z) = txt1
    End If
    SendKeys ("{down}")
    'Worksheets("TABLES").Cells(2, 2).Offset(Z, 0) = txt1
Next

SKP1:

u = 0
slow1
SendKeys ("%vdf"), True
slow1
SendKeys ("{Tab}")
CopyCompareCell
    If txt1 = "" Then BringToFrontT
    If txt1 = "" Then MsgBox "Rout Map out of alignment"
    If txt1 = "" Then Stop
For i = LBound(itemseq) To UBound(itemseq)
    CopyCompareCell
    If txt1 = "" Then GoTo SKP2
    If Not txt1 = "" Then
        u = u + 1
    ReDim Preserve RArray(0 To u - 1) ' Redimension:
        RArray(UBound(RArray)) = txt1 ' Fill last element
    End If
    SendKeys ("{down}")
Next
SKP2:

u = 0
slow1
SendKeys ("%vdf"), True
slow1
SendKeys ("{Tab}")
SendKeys ("{Tab}")
SendKeys ("{Tab}")
SendKeys ("{Tab}")
SendKeys ("{Tab}")
slow1
CopyCompareCell
    If txt1 = "" Then BringToFrontT
    If txt1 = "" Then MsgBox "Rout Map out of alignment"
    If txt1 = "" Then Stop
For i = LBound(itemseq) To UBound(itemseq)
    If txt1 = "" Then BringToFrontT
    If txt1 = "" Then MsgBox "Rout Map out of alignment"
    If txt1 = "" Then Stop
    CopyCompareCell
    If txt1 = "" Then GoTo SKP3
    If Not txt1 = "" Then
        u = u + 1
    ReDim Preserve IArray(0 To u - 1) ' Redimension:
        IArray(UBound(IArray)) = txt1 ' Fill last element
    End If
    SendKeys ("{down}")
Next

SKP3:

For i = LBound(itemseq) To UBound(itemseq)
    
    Cells(findCell.Row, 38) = Cells(findCell.Row, 38) & itemseq(i) & Chr(10)
    Cells(findCell.Row, 39) = Cells(findCell.Row, 39) & RArray(i) & Chr(10)
    Cells(findCell.Row, 40) = Cells(findCell.Row, 40) & IArray(i) & Chr(10)
    
Next

'updateROUT

'slow1
'enterNEWROUT

u = 0

SendKeys "+{PGUP}"
slow1
SendKeys ("%vf"), True
SendKeys ("{Tab}")
slow1
SendKeys ("SHLD RES"), True
SendKeys ("%i"), True
slow1
SendKeys ("%r"), True


slow1
Dim itemseq2() As Variant

Dim RArray2() As Variant
Dim IArray2() As Variant
slow1
CopyCompareCell
    If txt1 = "" Then BringToFrontT
    If txt1 = "" Then MsgBox "Rout Map out of alignment"
    If txt1 = "" Then Stop
For x = 1 To 5
    CopyCompareCell
    If txt1 = "" Then GoTo SKP4
    If Not txt1 = "" Then
    
    u = u + 1
    ReDim Preserve itemseq2(0 To x - 1) ' Redimension:
        itemseq2(UBound(itemseq2)) = txt1 ' Fill last element
    End If
    SendKeys ("{down}")
Next
SKP4:
u = 0
slow1
SendKeys ("%vdf"), True
SendKeys ("{Tab}")
CopyCompareCell
    If txt1 = "" Then BringToFrontT
    If txt1 = "" Then MsgBox "Rout Map out of alignment"
    If txt1 = "" Then Stop
For i = LBound(itemseq2) To UBound(itemseq2)
    CopyCompareCell
    If txt1 = "" Then BringToFrontT
    If txt1 = "" Then MsgBox "Rout Map out of alignment"
    If txt1 = "" Then Stop
    If Not txt1 = "" Then
        u = u + 1
    ReDim Preserve RArray2(0 To i) ' Redimension:
        RArray2(UBound(RArray2)) = txt1 ' Fill last element
    End If
    If Left(txt1, 5) = "PRESS" Then
        If Not (Cells(findCell.Row, 3)) = Empty Then
            SendKeys (Cells(findCell.Row, 3))
        End If
    End If
    SendKeys ("{down}")
Next

u = 0
slow1
SendKeys ("%vdf"), True
slow1
SendKeys ("{Tab}")
SendKeys ("{Tab}")
SendKeys ("{Tab}")
SendKeys ("{Tab}")
SendKeys ("{Tab}")
CopyCompareCell
    If txt1 = "" Then BringToFrontT
    If txt1 = "" Then MsgBox "Rout Map out of alignment"
    If txt1 = "" Then End
For i = LBound(itemseq2) To UBound(itemseq2)
    CopyCompareCell
    If txt1 = "" Then BringToFrontT
    If txt1 = "" Then MsgBox "Rout Map out of alignment"
    If txt1 = "" Then Stop
    If Not txt1 = "" Then
        u = u + 1
    ReDim Preserve IArray2(0 To i) ' Redimension:
        IArray2(UBound(IArray2)) = txt1 ' Fill last element
    End If
    SendKeys ("{down}")
Next

For i = LBound(itemseq2) To UBound(itemseq2)
    Cells(findCell.Row, 41) = Cells(findCell.Row, 41) & itemseq2(i) & Chr(10)
    Cells(findCell.Row, 42) = Cells(findCell.Row, 42) & RArray2(i) & Chr(10)
    Cells(findCell.Row, 43) = Cells(findCell.Row, 43) & IArray2(i) & Chr(10)
Next
slow1
SendKeys "%vdi"

SendKeys ("{Tab}")
CopyCompareCell
If Not txt1 = Cells(findCell.Row, 3) Then
    MsgBox "Tool number doesn't match, please enter changes to SHLD RES manually, exit routing window and then re-start Scirpt"
    Cells(findCell.Row, 30) = "mapBOM"
    'End
End If
slow1
SendKeys ("{Tab}")
SendKeys ("{Tab}")
SendKeys ("{Tab}")
SendKeys ("{Tab}")
slow1
For i = LBound(itemseq2) To UBound(itemseq2)
    SendKeys Cells(findCell.Row, 9)
    SendKeys ("{down}")
    slow1
Next




'find press seq, and update it
'if no press is present then add it




slow2
SendKeys "%fs"
slow1
SendKeys "%fc"


Cells(findCell.Row, 30).Value = "RunBOM"



SkipToEnd:

End Sub
Sub updateSCHD()

SendKeys ("{Tab}")
CopyCompareCell
If Not txt1 = Cells(findCell.Row, 3) Then
    MsgBox "Tool number doesn't match, please enter changes to SHLD RES manually, exit routing window and then re-start Scirpt"
    Cells(findCell.Row, 30) = "mapBOM"
    End
End If
slow1
SendKeys ("{Tab}")
SendKeys ("{Tab}")
SendKeys ("{Tab}")
SendKeys ("{Tab}")
slow1
SendKeys Cells(findCell.Row, 9)


End Sub
Public Function GetLength(a As Variant) As Integer
   If IsEmpty(a) Then
      GetLength = 0
   Else
      GetLength = UBound(a) - LBound(a) + 1
   End If
End Function
Sub BomMap()
Dim u As Integer
u = 0
Dim purArray() As Variant
Dim Cu As Range
Dim wrRng As Range
Set wrRng = ThisWorkbook.Worksheets("CompCostUpdate").Range("A3:A213")



ThisWorkbook.Worksheets("Transfer Line Calculator").Activate
Application.FindFormat.Clear
Application.FindFormat.Interior.ColorIndex = 2

Dim rng As Range
Set rng = ThisWorkbook.Worksheets("Transfer Line Calculator").Range("B7:B113")
Set findCell = rng.Find(what:="*", searchFormat:=True)
Cells(findCell.Row, 44) = ""
Cells(findCell.Row, 45) = ""
slow1
ClickOnCornerWindow
setvars

For Each Cu In wrRng

    If Cu.Offset(0, 6) = "PURCHASED" Then
        ReDim Preserve purArray(0 To u) ' Redimension:
        purArray(UBound(purArray)) = Cu.Value ' Fill last element
        u = u + 1
        
    End If

Next Cu

'If purArray = Empty Then purArray(0) = k

ActiveWorkbook.Worksheets("TABLES").Range("B1:f30") = ""
slow2
SendKeys ("%fw"), True

slow1
SendKeys ("nb"), True
slow1


changeBM
SendKeys ("%tc"), True

slow1
SendKeys ("%tu"), True
slow1
SendKeys ("b"), True
SendKeys ("~")

'enter item and check 00 to ensure we don't encouter item alread exsist error
slow1
SendKeys ("%vf"), True
slow1
SendKeys ItemNum
SendKeys ("~")
slow2
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
slow1
txt1 = Empty

Dim txt2 As String
Dim txt3 As String
Dim txt4 As String
txt2 = DateValue(WorksheetFunction.Text(Date, "mm/dd/yyyy"))
Dim ClipObj1 As New DataObject
Application.SendKeys ("^c")
DoEvents
Application.Wait (Now + TimeValue("00:00:01"))

ClipObj1.GetFromClipboard
On Error Resume Next
    txt4 = ClipObj1.GetText(1)
    txt3 = DateValue(WorksheetFunction.Text(txt4, "mm/dd/yyyy"))
    If Not txt3 = txt2 Then Stop
    If Not txt3 = txt2 Then EndBOM
On Error GoTo 0
ClearClipboard



slow1
SendKeys "+{PGDN}"
SendKeys "%vdl"
Dim itemseqz() As Variant

u = 0
Dim RArray() As Variant
Dim OArray() As Variant
Dim UArray() As Variant
Dim LoArray() As Variant
Dim SiArray() As Variant
slow1
CopyCompareCell

    If txt1 = "" Then BringToFrontT
    If txt1 = "" Then MsgBox "Rout Map out of alignment"
    If txt1 = "" Then Stop
    
'create temp map of Bom that is in Oracle
For Z = 0 To 20
    CopyCompareCell
    If txt1 = "" Then ReDim Preserve itemseqz(0 To Z - 1)
    If Not Z = 0 Then
        If itemseqz(UBound(itemseqz)) = txt1 Then GoTo SKP1
    End If
    If Not txt1 = "" Then
    u = u + 1
    ReDim Preserve itemseqz(0 To Z) ' Redimension:
    itemseqz(UBound(itemseqz)) = txt1 ' Fill last element
    'itemseqz(Z) = txt1
    End If
    SendKeys ("{up}")
    Worksheets("TABLES").Cells(2, 2).Offset(Z, 0) = txt1
Next

SKP1:

'for each cell that isn't empty, add to transferline box

Z = 0
slow1
Dim SeqY() As Variant
'create temp array of the Seq #s used
For Z = 0 To 20

  If Not Worksheets("TABLES").Cells(2, 2).Offset(Z, 0) = "" Then
        ReDim Preserve SeqY(0 To Z) ' Redimension:
        SeqY(UBound(SeqY)) = Worksheets("TABLES").Cells(2, 2).Offset(Z, 0) ' Fill last element
  End If
    
Next



u = 0
slow1
SendKeys ("%vdl"), True
slow1
SendKeys ("{Tab}")
SendKeys ("{Tab}")
CopyCompareCell
    If txt1 = "" Then BringToFrontT
    If txt1 = "" Then MsgBox "Rout Map out of alignment"
    If txt1 = "" Then Stop
    
'create temp array of old Item #s used
For i = LBound(itemseqz) To UBound(itemseqz)
    CopyCompareCell
    If txt1 = "" Then GoTo SKP2
    If Not txt1 = "" Then
        u = u + 1
    ReDim Preserve RArray(0 To u - 1) ' Redimension:
        RArray(UBound(RArray)) = txt1 ' Fill last element
    End If
    SendKeys ("{up}")
Next
SKP2:

For i = UBound(itemseqz) To LBound(itemseqz) Step -1

    Cells(findCell.Row, 44) = Cells(findCell.Row, 44) & itemseqz(i) & Chr(10)
    Cells(findCell.Row, 45) = Cells(findCell.Row, 45) & RArray(i) & Chr(10)
    
Next

'new item array
Cells(1, 44) = Cells(findCell.Row, 44)
Cells(1, 45) = Cells(findCell.Row, 45)
Cells(1, 50) = WorksheetFunction.Concat(Cells(findCell.Row, 20), Chr(10), Cells(findCell.Row, 26), Chr(10), Cells(findCell.Row, 24), Chr(10))
Cells(1, 51) = WorksheetFunction.Concat(Cells(findCell.Row, 21), Chr(10), Cells(findCell.Row, 27), Chr(10), Cells(findCell.Row, 25), Chr(10))



For p = 0 To 20
    
    ReDim Preserve OArray(0 To p) ' Redimension:
    ReDim Preserve UArray(0 To p)
    OArray(UBound(OArray)) = Left(Cells(1, 50), InStr(1, Cells(1, 50), Chr(10)) - 1)
    UArray(UBound(UArray)) = Left(Cells(1, 51), InStr(1, Cells(1, 51), Chr(10)) - 1)
    Cells(1, 50) = Mid(Cells(1, 50), InStr(1, Cells(1, 50), Chr(10)) + 1)
    Cells(1, 51) = Mid(Cells(1, 51), InStr(1, Cells(1, 51), Chr(10)) + 1)
    If Cells(1, 50) = "" Then GoTo skpp

Next
skpp:

For e = 1 To UBound(Split(Cells(findCell.Row, 20), Chr(10)))
    
     Cells(1, 52) = Cells(1, 52) & " " & Chr(10)
     Cells(1, 53) = Cells(1, 53) & " " & Chr(10)
Next
If Not Cells(findCell.Row, 28) = "" Then
    Cells(1, 52) = Cells(1, 52) & Chr(10) & Cells(findCell.Row, 28)
    Cells(1, 53) = Cells(1, 53) & Chr(10) & Cells(findCell.Row, 29)
End If

Cells(1, 52) = Cells(1, 52) & Chr(10)
Cells(1, 53) = Cells(1, 53) & Chr(10)
e = 0

For e = 0 To (UBound(Split(Cells(findCell.Row, 26), Chr(10))) + UBound(Split(Cells(findCell.Row, 20), Chr(10))) + 1)
    If e < UBound(Split(Cells(findCell.Row, 26), Chr(10))) + 1 Then
        ReDim Preserve LoArray(0 To e) ' Redimension:
        ReDim Preserve SiArray(0 To e)
        LoArray(UBound(LoArray)) = Left(Cells(1, 52), InStr(1, Cells(1, 52), Chr(10)) - 1)
        SiArray(UBound(SiArray)) = Left(Cells(1, 53), InStr(1, Cells(1, 53), Chr(10)) - 1)
        Cells(1, 52) = Mid(Cells(1, 52), InStr(1, Cells(1, 52), Chr(10)) + 1)
        Cells(1, 53) = Mid(Cells(1, 53), InStr(1, Cells(1, 53), Chr(10)) + 1)
        
    Else
        
        ReDim Preserve LoArray(0 To e) ' Redimension:
        ReDim Preserve SiArray(0 To e)
        LoArray(UBound(LoArray)) = Left(Cells(1, 52), InStr(1, Cells(1, 52), Chr(10)) - 1)
        SiArray(UBound(SiArray)) = Left(Cells(1, 53), InStr(1, Cells(1, 53), Chr(10)) - 1)
        Cells(1, 52) = Mid(Cells(1, 52), InStr(1, Cells(1, 52), Chr(10)) + 1)
        Cells(1, 53) = Mid(Cells(1, 53), InStr(1, Cells(1, 53), Chr(10)) + 1)
        If Cells(1, 52) = "" Then GoTo skpm
    End If


Next
skpm:

Worksheets("TABLES").Range("B2:B20") = ""

Dim ItemMisMatch As Boolean
Dim Itemfound As Boolean
Dim seqM As Integer
Dim MSGX As String
MSGX = ""
seqM = 10
' loop of searching old bom items in new bom notifying user if it needs intervention to remove item from BOM.
Cells(1, 44) = ""
Cells(1, 45) = ""
For Z = LBound(RArray) To UBound(RArray)

    For Y = LBound(OArray) To UBound(OArray)
        If Y > UBound(OArray) Then GoTo N1
        If OArray(Y) = RArray(Z) Then Itemfound = True
        If OArray(Y) = RArray(Z) Then Cells(1, 46) = Cells(1, 46) & UArray(Y) & Chr(10)
    Next
If SeqY(Z) > seqM Then seqM = SeqY(Z)
If Itemfound = True Then Cells(1, 44) = Cells(1, 44) & SeqY(Z) & Chr(10)
If Itemfound = True Then Cells(1, 45) = Cells(1, 45) & RArray(Z) & Chr(10)
'If Itemfound = True Then Cells(1, 46) = Cells(1, 46) & UArray(Z) & Chr(10)
If Itemfound = True Then Worksheets("TABLES").Cells(2, 2).Offset(Z + 1, 0) = SeqY(Z)
If Itemfound = True Then Worksheets("TABLES").Cells(2, 3).Offset(Z + 1, 0) = RArray(Z)

N1:
'If Itemfound = False Then BringToFrontT
If Itemfound = False Then MSGX = MSGX & "Exsisting Item not found " & RArray(Z) & " in New Item List" & Chr(10)
'enter request to end date items
If Itemfound = False Then RArray(Z) = ""
If Itemfound = False Then SeqY(Z) = ""
'If Itemfound = False Then Stop
Itemfound = False

Next





If Not MSGX = "" Then MSGX = MSGX & Chr(10) & "Please end Date old items and re-start script"
If Not MSGX = "" Then BringToFrontT
If Not MSGX = "" Then MsgBox MSGX
If Not MSGX = "" Then Cells(findCell.Row, 30) = "ItemsENDED"
Dim t As Integer
Dim ct As Integer
ct = 1
' loop of searching old bom items in new bom and assigning used & new Seq values to new Bom Map.
For t = LBound(OArray) To UBound(OArray)
Itemfound = False

    For k = LBound(RArray) To UBound(RArray)
    If k > UBound(OArray) Then GoTo P2
        If RArray(k) = OArray(t) Then Itemfound = True
    Next
If Itemfound = False Then Cells(1, 44) = Cells(1, 44) & seqM + 10 * ct & Chr(10)
If Itemfound = False Then seqM = seqM + 10
If Itemfound = False Then Cells(1, 45) = Cells(1, 45) & OArray(t) & Chr(10)
If Itemfound = False Then Cells(1, 46) = Cells(1, 46) & UArray(t) & Chr(10)
On Error Resume Next
If Itemfound = False Then Cells(1, 47) = Cells(1, 47) & LoArray(t) & Chr(10)
If Itemfound = False Then Cells(1, 48) = Cells(1, 48) & SiArray(t) & Chr(10)
On Error GoTo 0
If Itemfound = False Then Worksheets("TABLES").Cells(2, 2).Offset(Z + t + 1, 0) = seqM
If Itemfound = False Then Worksheets("TABLES").Cells(2, 3).Offset(Z + t + 1, 0) = OArray(t)
If Itemfound = False Then Worksheets("TABLES").Cells(2, 4).Offset(Z + t + 1, 0) = UArray(t)
On Error Resume Next
If Itemfound = False Then Worksheets("TABLES").Cells(2, 5).Offset(Z + t + 1, 0) = LoArray(t)
If Itemfound = False Then Worksheets("TABLES").Cells(2, 6).Offset(Z + t + 1, 0) = SiArray(t)
On Error GoTo 0

Next
P2:

'sort small to large seqz section in TABLES Sheet

With ActiveWorkbook.Worksheets("TABLES").Sort
 .SortFields.Clear
 .SortFields.Add Key:=Range("B1:B30"), _
 SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
 .SetRange Range("B1:f30")
 .Header = xlYes
 .MatchCase = False
 .Orientation = xlTopToBottom
 .SortMethod = xlPinYin
 .Apply
End With


Cells(findCell.Row, 44) = ""
Cells(findCell.Row, 45) = ""

Dim RY As Range
For Each RY In Worksheets("TABLES").Range("B1:B30")
    
    If RY = "" Then GoTo brk
    Cells(findCell.Row, 44) = Cells(findCell.Row, 44) & RY & Chr(10)
    Cells(findCell.Row, 45) = Cells(findCell.Row, 45) & RY.Offset(0, 1) & Chr(10)
    Cells(findCell.Row, 46) = Cells(findCell.Row, 46) & RY.Offset(0, 2) & Chr(10)
    Cells(findCell.Row, 47) = Cells(findCell.Row, 47) & RY.Offset(0, 3) & Chr(10)
    Cells(findCell.Row, 48) = Cells(findCell.Row, 48) & RY.Offset(0, 4) & Chr(10)
brk:

Next



Cells(1, 44) = ""
Cells(1, 45) = ""
Cells(1, 46) = ""
Cells(1, 47) = ""
Cells(1, 48) = ""
' create temp BOM with Use values.
' add use amounts to tables, then clear feilds and re-load values


Cells(findCell.Row, 30) = "BOMreMaped"

'End

SendKeys "%fc"
UpdateBOM
Stop
End Sub
Sub UpdateBOM()


ThisWorkbook.Worksheets("Transfer Line Calculator").Activate
Application.FindFormat.Clear
Application.FindFormat.Interior.ColorIndex = 2

Dim rng As Range
Set rng = ThisWorkbook.Worksheets("Transfer Line Calculator").Range("B7:B113")
Set findCell = rng.Find(what:="*", searchFormat:=True)
Cells(findCell.Row, 44) = ""
Cells(findCell.Row, 45) = ""
slow1
ClickOnCornerWindow
setvars

slow2
SendKeys ("%fw"), True

slow1
SendKeys ("nb"), True
slow1


changeBM
SendKeys ("%tc"), True

slow1
SendKeys ("%tu"), True
slow1
SendKeys ("b"), True
SendKeys ("~")

'enter item and check 00 to ensure we don't encouter item alread exsist error
slow1
SendKeys ("%vf"), True
slow1
SendKeys ItemNum
SendKeys ("~")
slow2
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
slow1
txt1 = Empty

Dim txt2 As String
Dim txt3 As String
Dim txt4 As String
txt2 = DateValue(WorksheetFunction.Text(Date, "mm/dd/yyyy"))
Dim ClipObj1 As New DataObject
Application.SendKeys ("^c")
DoEvents
Application.Wait (Now + TimeValue("00:00:01"))

ClipObj1.GetFromClipboard
On Error Resume Next
    txt4 = ClipObj1.GetText(1)
    txt3 = DateValue(WorksheetFunction.Text(txt4, "mm/dd/yyyy"))
    If Not txt3 = txt2 Then Stop
    If Not txt3 = txt2 Then EndBOM
On Error GoTo 0
ClearClipboard

slow1
SendKeys "+{PGDN}"


Dim RNG2 As Range
Dim RP As Range
Set RNG2 = Worksheets("TABLES").Range("B2:B30")

    For Each RP In RNG2
    
        On Error Resume Next
        If IsInArray(RP.Offset(0, 1).Value, purArray) = True Then
            If Cells(findCell.Row, 22) = "SLB" Then
                RP.Offset(0, 3) = "U40 ASSY"
                RP.Offset(0, 4) = "MAN.00.00"
            End If
        End If
        On Error GoTo 0
        If Left(RP.Offset(0, 2), 1) = "9" Then
            If Cells(findCell.Row, 22) = "LVG" Then
                RP.Offset(0, 3) = "U05 RMM"
                RP.Offset(0, 4) = "WM05.01.00"
            End If
        End If
        
        If Left(RP.Offset(0, 1), 1) = "9" Then resinLoad
        If Left(RP.Offset(0, 1), 1) = "R" Then packLoad
        If Left(RP.Offset(0, 1), 1) = "N" Then packLoad
        If Left(RP.Offset(0, 1), 1) = "B" Then packLoad
        If Left(RP.Offset(0, 1), 1) = "X" Then packLoad
        If Left(RP.Offset(0, 1), 1) >= 0 And Left(RP.Offset(0, 1), 1) < 9 Then compLoad
       
    Next RP

Stop
End Sub
Sub resinLoad()

CopyCompareCell
If Not RP = txt1 Then BringToFrontT
If Not RP = txt1 Then MsgBox "SEQ does not match, Out of alignment, stopting script"
If Not RP = txt1 Then Stop

SendKeys "{tab}"
SendKeys "{tab}"
SendKeys RP.Offset(0, 1)
For i = 1 To 7
        Application.SendKeys ("{Tab}")
        slow1
        CopyCompareCell
        'txt1 = ClipObj1.GetText(1)
        On Error Resume Next
        txt1 = DateValue(WorksheetFunction.Text(txt1, "mm/dd/yyyy"))
        On Error GoTo 0
        If txt1 = txt3 Then i = 8
        
Next
Application.SendKeys ("+{Tab}")
Application.SendKeys ("+{Tab}")
Application.SendKeys ("+{Tab}")

Application.SendKeys RP.Offset(0, 2)
slow1
For r = 1 To 20
    slow1
    CopyCompareCell
    
    If txt1 = "100" Then r = 21
    If txt1 = "100" Then GoTo ENDFOR
    
    slow1
    Application.SendKeys ("{Tab}")
    
    slow1
ENDFOR:
Next

For r = 1 To 20
    slow1
    CopyCompareCell
    If txt1 = "1" Then r = 21
    If txt1 = "1" Then Application.SendKeys ".99"
    If txt1 = ".99" Then r = 21
    If txt1 = ".99" Then GoTo ENDFOR1
    slow1
    If txt1 = "1" Then GoTo ENDFOR1
    slow1
    Application.SendKeys ("{Tab}")
    
    slow1
ENDFOR1:
Next
 If Not txt1 = "Assembly Pull" Then
        Application.SendKeys ("{Tab}")
        slow1
        CopyCompareCell
        If Not txt1 = "Assembly Pull" Then
            Application.SendKeys ("{Tab}")
            slow1
            CopyCompareCell
            If Not txt1 = "Assembly Pull" Then
                Application.SendKeys ("{Tab}")
                slow1
                CopyCompareCell
                If Not txt1 = "Assembly Pull" Then
                Application.SendKeys ("{Tab}")
                slow1
                CopyCompareCell
                End If
            End If
        End If
    End If

        slow1
        
        If Cells(findCell.Row, 22).Value = "CNL" Then
            If i = 1 Then
               'CopyCompareCell
                Application.SendKeys "Push"
                Else
                Application.SendKeys "Assembly Pull"
            End If
        End If
        
        
        CopyCompareCell
        If Cells(findCell.Row, 22).Value = "SLB" Then
            If txt1 = "Assembly Pull" Then
                Application.SendKeys ("{Tab}")
                Application.SendKeys "SLB PROD"
            End If
        End If

Application.SendKeys ("{Tab}")
If Not RP.Offset(0, 3) Then Application.SendKeys RP.Offset(0, 3)
Application.SendKeys ("{Tab}")
If Not RP.Offset(0, 4) Then Application.SendKeys RP.Offset(0, 4)


slow1
Application.SendKeys ("{down}")
End If


End Sub
Sub packLoad()

CopyCompareCell
If Not RP = txt1 Then BringToFrontT
If Not RP = txt1 Then MsgBox "SEQ does not match, Out of alignment, stopting script"
If Not RP = txt1 Then Stop

SendKeys "{tab}"
SendKeys "{tab}"
SendKeys RP.Offset(0, 1)
For i = 1 To 7
        Application.SendKeys ("{Tab}")
        slow1
        CopyCompareCell
        'txt1 = ClipObj1.GetText(1)
        On Error Resume Next
        txt1 = DateValue(WorksheetFunction.Text(txt1, "mm/dd/yyyy"))
        On Error GoTo 0
        If txt1 = txt3 Then i = 8
        
Next

Application.SendKeys ("+{Tab}")
slow1
Application.SendKeys ("+{Tab}")
slow1
Application.SendKeys RP.Offset(0, 2)
   
    
    slow1
    'Application.SendKeys ("{Tab}")
    slow1
       If Left(RP.Offset(0, 1), 1) = "R" Then
        Application.SendKeys ("{Tab}")
        Application.SendKeys ("{Tab}")
        Application.SendKeys ("{Tab}")
        slow1
        Application.SendKeys ("{Tab}")
        Application.SendKeys ("{Tab}")
        Application.SendKeys ("{Tab}")
        slow1
        Application.SendKeys ("{Tab}")
        Application.SendKeys ("{Tab}")
        Application.SendKeys ("{Tab}")
        slow1
        Application.SendKeys "Bulk"
        Else
        Application.SendKeys ("{Tab}")
        Application.SendKeys ("{Tab}")
        Application.SendKeys ("{Tab}")
        slow1
        Application.SendKeys ("{Tab}")
        Application.SendKeys ("{Tab}")
        Application.SendKeys ("{Tab}")
        slow1
        Application.SendKeys ("{Tab}")
         CopyCompareCell
        If Not txt1 = "Assembly Pull" Then
            Application.SendKeys ("{Tab}")
            slow1
            CopyCompareCell
            If Not txt1 = "Assembly Pull" Then
                Application.SendKeys ("{Tab}")
                slow1
                CopyCompareCell
                If Not txt1 = "Assembly Pull" Then
                    Application.SendKeys ("{Tab}")
                    slow1
                    CopyCompareCell
                        If Not txt1 = "Assembly Pull" Then
                            Application.SendKeys ("{Tab}")
                            slow1
                            CopyCompareCell
                        End If
                End If
            End If
        End If
            If Cells(findCell.Row, 22) = "SLB" Then
            Application.SendKeys ("{Tab}")
            slow1
            Application.SendKeys "SLB PROD"
        End If
    End If
    slow1
Application.SendKeys ("{Tab}")
If Not RP.Offset(0, 3) Then Application.SendKeys RP.Offset(0, 3)
Application.SendKeys ("{Tab}")
If Not RP.Offset(0, 4) Then Application.SendKeys RP.Offset(0, 4)

    
Application.SendKeys ("{down}")

End Sub
Sub compLoad()

CopyCompareCell
If Not RP = txt1 Then BringToFrontT
If Not RP = txt1 Then MsgBox "SEQ does not match, Out of alignment, stopting script"
If Not RP = txt1 Then Stop

SendKeys "{tab}"
SendKeys "{tab}"
SendKeys RP.Offset(0, 1)
For i = 1 To 7
        Application.SendKeys ("{Tab}")
        slow1
        CopyCompareCell
        'txt1 = ClipObj1.GetText(1)
        On Error Resume Next
        txt1 = DateValue(WorksheetFunction.Text(txt1, "mm/dd/yyyy"))
        On Error GoTo 0
        If txt1 = txt3 Then i = 8
        
Next

Application.SendKeys ("+{Tab}")
slow1
Application.SendKeys ("+{Tab}")
slow1
Application.SendKeys ("+{Tab}")
slow1
Application.SendKeys RP.Offset(0, 2)
   

slow1
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
slow1
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
slow1
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
slow1
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
'we are at None
slow2
CopyCompareCell
If Not txt1 = "Assembly Pull" Then
    Application.SendKeys ("{Tab}")
    slow1
    CopyCompareCell
    If Not txt1 = "Assembly Pull" Then
        Application.SendKeys ("{Tab}")
        slow1
        CopyCompareCell
        If Not txt1 = "Assembly Pull" Then
            Application.SendKeys ("{Tab}")
            slow1
            CopyCompareCell
        End If
    End If
End If
slow2

    'Application.SendKeys ("{Tab}")
    'Application.SendKeys "Assembly Pull"
Application.SendKeys ("{Tab}")
If Not RP.Offset(0, 3) Then Application.SendKeys RP.Offset(0, 3)
Application.SendKeys ("{Tab}")
If Not RP.Offset(0, 4) Then Application.SendKeys RP.Offset(0, 4)

slow1
Application.SendKeys ("{down}")



End Sub

Sub ReplaceCost()



Application.FindFormat.Clear
Application.FindFormat.Interior.ColorIndex = 2

Dim rng As Range
Set rng = ThisWorkbook.Worksheets("Transfer Line Calculator").Range("B7:B113")
Set findCell = rng.Find(what:="*", searchFormat:=True)

slow1

'setvars

slow2
SendKeys ("%fw"), True

slow1
SendKeys ("cc"), True
slow1


changeOrg

SendKeys ("%to"), True

SendKeys ("~")
SendKeys ("~")
slow1
SendKeys ItemNum
SendKeys ("{Tab}")
SendKeys Cells(findCell.Row, 36).Value
SendKeys ("%i")
slow1
SendKeys ("%c")
slow1
slow1
SendKeys ("%err")

If Cells(findCell.Row, 4) = "Shoot & Ship" Then enterShootShip
If Cells(findCell.Row, 4) = "Molded Component" Then enterMoldedComp
If Cells(findCell.Row, 4) = "Assembly" Then enterAssembly
If Cells(findCell.Row, 4) = "Sub Assembly" Then enterAssembly



slow1
SendKeys ("%fs")
slow1
slow1
SendKeys ("%fc")
findCell.Interior.ColorIndex = 6

End Sub
Sub enterAssembly()
'sub for entering cost elements, based on item type, assembly

slow2
Application.SendKeys ("Material Overhead")
Application.SendKeys ("{Tab}")
Application.SendKeys ("Yield ")
Application.SendKeys Cells(findCell.Row, 23).Value
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
slow1
If Cells(findCell.Row, 4) = "Sub Assembly" Then Application.SendKeys ".02"
If Cells(findCell.Row, 4) = "Assembly" Then Application.SendKeys ".01"
slow1

If Cells(findCell.Row, 34) = "Yes" Then GoTo DSPSITEM
Application.SendKeys ("{down}")
Application.SendKeys ("Material Overhead")
Application.SendKeys ("{Tab}")
Application.SendKeys ("Admin")
Application.SendKeys Mid(Cells(findCell.Row, 23).Value, 2, 5)
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
slow1
If Cells(findCell.Row, 4) = "Sub Assembly" Then Application.SendKeys ".02"
If Cells(findCell.Row, 4) = "Assembly" Then Application.SendKeys ".06"
slow1

Application.SendKeys ("{down}")
Application.SendKeys ("Material Overhead")
Application.SendKeys ("{Tab}")
Application.SendKeys ("TechCo")
Application.SendKeys Mid(Cells(findCell.Row, 23).Value, 2, 5)
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
slow1
Application.SendKeys ".04"
slow1
DSPSITEM:
If Cells(findCell.Row, 4) = "Sub Assembly" Then GoTo SkipRoyty

Application.SendKeys ("{down}")
Application.SendKeys ("Material Overhead")
Application.SendKeys ("{Tab}")
Application.SendKeys ("Rylty. ")
Application.SendKeys Cells(findCell.Row, 23).Value
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
slow1
Application.SendKeys Cells(findCell.Row, 15).Value
slow1
SkipRoyty:
If Not Cells(findCell.Row, 11) = "" Then
    Application.SendKeys ("{down}")
    Application.SendKeys ("Material Overhead")
    Application.SendKeys ("{Tab}")
    Application.SendKeys ("R-MEX FRGH")
    Application.SendKeys ("{Tab}")
    Application.SendKeys ("{Tab}")
    Application.SendKeys ("{Tab}")
    slow1
    Application.SendKeys Cells(findCell.Row, 11).Value
    slow1
End If
    
End Sub
Sub enterShootShip()
'sub for entering cost elements, based on item type, shoot and ship

GrabToolCost

slow2
Application.SendKeys ("Material Overhead")
Application.SendKeys ("{Tab}")
Application.SendKeys ("Yield ")
Application.SendKeys Cells(findCell.Row, 23).Value
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
slow1
Application.SendKeys ".01"
slow2

Application.SendKeys ("{down}")
Application.SendKeys ("Material Overhead")
Application.SendKeys ("{Tab}")
Application.SendKeys ("AuxCo")
Application.SendKeys Mid(Cells(findCell.Row, 23).Value, 2, 5)
slow1
If Cells(findCell.Row, 23) = "U18" Then
    Application.SendKeys ("{Tab}")
    slow1
    SendKeys ("{down}")
    Application.SendKeys "%o"
    slow1
    
    Application.SendKeys ("+{Tab}")
End If

Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
slow1
Application.SendKeys ".14"
slow1

Application.SendKeys ("{down}")
Application.SendKeys ("Material Overhead")
Application.SendKeys ("{Tab}")
Application.SendKeys ("TechCo")
Application.SendKeys Mid(Cells(findCell.Row, 23).Value, 2, 5)
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
slow1
Application.SendKeys ".06"
slow1

Application.SendKeys ("{down}")
Application.SendKeys ("Material Overhead")
Application.SendKeys ("{Tab}")
Application.SendKeys ("ToolRp")
Application.SendKeys Mid(Cells(findCell.Row, 23).Value, 2, 5)
Application.SendKeys ("-")
Application.SendKeys Cells(findCell.Row, 17).Value
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
slow1
Application.SendKeys ToolCost
slow1

Application.SendKeys ("{down}")
Application.SendKeys ("Material Overhead")
Application.SendKeys ("{Tab}")
Application.SendKeys ("Rylty. ")
Application.SendKeys Cells(findCell.Row, 23).Value
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
slow1
Application.SendKeys Cells(findCell.Row, 15).Value
slow1

If Not Cells(findCell.Row, 11) = "" Then
    Application.SendKeys ("{down}")
    Application.SendKeys ("Material Overhead")
    Application.SendKeys ("{Tab}")
    Application.SendKeys ("R-MEX FRGH")
    Application.SendKeys ("{Tab}")
    Application.SendKeys ("{Tab}")
    Application.SendKeys ("{Tab}")
    slow1
    Application.SendKeys Cells(findCell.Row, 11).Value
    slow1
End If
If Not Cells(findCell.Row, 13) = "" Then
    Application.SendKeys ("{down}")
    Application.SendKeys ("Material Overhead")
    Application.SendKeys ("{Tab}")
    Application.SendKeys ("Glass")
    Application.SendKeys Mid(Cells(findCell.Row, 23).Value, 2, 5)
    Application.SendKeys ("{Tab}")
    Application.SendKeys ("{Tab}")
    Application.SendKeys ("{Tab}")
    slow1
    Application.SendKeys Cells(findCell.Row, 13).Value
    slow1
End If

If Not Cells(findCell.Row, 18) = "Yes" Then
    Application.SendKeys ("{down}")
    Application.SendKeys ("Material")
    Application.SendKeys ("{Tab}")
    Application.SendKeys ("MTL REGRND")
    Application.SendKeys ("{Tab}")
    Application.SendKeys ("{Tab}")
    Application.SendKeys ("{Tab}")
    slow1
    Application.SendKeys Cells(findCell.Row, 14).Value
    slow1
End If

    
End Sub
Sub enterMoldedComp()
'sub for entering cost elements, based on item type, molded

GrabToolCost

slow2
Application.SendKeys ("Material Overhead")
Application.SendKeys ("{Tab}")
Application.SendKeys ("Yield ")
Application.SendKeys Cells(findCell.Row, 23).Value
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
slow1
Application.SendKeys ".01"
slow1

Application.SendKeys ("{down}")
Application.SendKeys ("Material Overhead")
Application.SendKeys ("{Tab}")
Application.SendKeys ("AuxCo")
Application.SendKeys Mid(Cells(findCell.Row, 23).Value, 2, 5)
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
slow1
Application.SendKeys ".14"
slow1

Application.SendKeys ("{down}")
Application.SendKeys ("Material Overhead")
Application.SendKeys ("{Tab}")
Application.SendKeys ("TechCo")
Application.SendKeys Mid(Cells(findCell.Row, 23).Value, 2, 5)
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
slow1
Application.SendKeys ".06"
slow1

Application.SendKeys ("{down}")
Application.SendKeys ("Material Overhead")
Application.SendKeys ("{Tab}")
Application.SendKeys ("ToolRp")
Application.SendKeys Mid(Cells(findCell.Row, 23).Value, 2, 5)
Application.SendKeys ("-")
Application.SendKeys Cells(findCell.Row, 17).Value
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
slow1
Application.SendKeys ToolCost
slow1

If Not Cells(findCell.Row, 18) = True Then
    Application.SendKeys ("{down}")
    Application.SendKeys ("Material")
    Application.SendKeys ("{Tab}")
    Application.SendKeys ("MTL REGRND")
    Application.SendKeys ("{Tab}")
    Application.SendKeys ("{Tab}")
    Application.SendKeys ("{Tab}")
    slow1
    Application.SendKeys Cells(findCell.Row, 14).Value
    slow1
End If
If Not Cells(findCell.Row, 13) = "" Then
    Application.SendKeys ("{down}")
    Application.SendKeys ("Material Overhead")
    Application.SendKeys ("{Tab}")
    Application.SendKeys ("Glass")
    Application.SendKeys Mid(Cells(findCell.Row, 23).Value, 2, 5)
    Application.SendKeys ("{Tab}")
    Application.SendKeys ("{Tab}")
    Application.SendKeys ("{Tab}")
    slow1
    Application.SendKeys Cells(findCell.Row, 13).Value
    slow1
End If
    
End Sub
Sub ResetBomForm()
'this sub sets and re-sets the add item form for kickoffs
'Dim iRow As Long

'iRow = [Counta(Kickoff Boms!B:B)]

    With TransferForm
        
        'mandatory field
        .ComboBox11.Clear
        .ComboBox11.AddItem "CNL"
        .ComboBox11.AddItem "GWH"
        .ComboBox11.AddItem "LVG"
        .ComboBox11.AddItem "MEX"
        .ComboBox11.AddItem "SLB"

        .ComboBox10.Clear
        .ComboBox10.AddItem "WHSE SAMPL"
        .ComboBox10.AddItem "CP WIP"
        .ComboBox10.AddItem "U20 COMPT"
        .ComboBox10.AddItem "CNL NOHOME"
        .ComboBox10.AddItem "CNL MAINT"
        .ComboBox10.AddItem "CNL NC"
        .ComboBox10.AddItem "CNL SHIPFG"
        .ComboBox10.AddItem "CNL TOOL"
        .ComboBox10.AddItem "FG STAGE"
        .ComboBox10.AddItem "MATERIAL"
        .ComboBox10.AddItem "PACKAGING"
        .ComboBox10.AddItem "U07 FG"
        .ComboBox10.AddItem "CNL WHSE"
        .ComboBox10.AddItem "GRV WHSE"
        .ComboBox10.AddItem "INBREPACK"
        .ComboBox10.AddItem "CNL REC"
        .ComboBox10.AddItem "U20 WIP"
        .ComboBox10.AddItem "PT REC"
        .ComboBox10.AddItem "U01 WIP"
        .ComboBox10.AddItem "PROJECT"
        .ComboBox10.AddItem "IRAPUATO"
        .ComboBox10.AddItem "CNL U01 NC"
        .ComboBox10.AddItem "U18NC"
        .ComboBox10.AddItem "NIPPON WHS"
        .ComboBox10.AddItem "NIPPON REC"
        .ComboBox10.AddItem "NIPPONSHIP"
        .ComboBox10.AddItem "CNL PACK"
        .ComboBox10.AddItem "CNLPRJSAFE"
        .ComboBox10.AddItem "GWH REC"
        .ComboBox10.AddItem "GWH SHIPFG"
        .ComboBox10.AddItem "WHSE SAMPL"
        .ComboBox10.AddItem "GWH TOOL"
        .ComboBox10.AddItem "U20STOCK"
        .ComboBox10.AddItem "CWH WHSE"
        .ComboBox10.AddItem "CWH REC"
        .ComboBox10.AddItem "CWH SHIP"
        .ComboBox10.AddItem "CWH PACK"
        .ComboBox10.AddItem "CWH U20"
        .ComboBox10.AddItem "U20 NC"
        .ComboBox10.AddItem "CWH ENDCAP"
        .ComboBox10.AddItem "CWH LOG NC"
        .ComboBox10.AddItem "INRPK"
        .ComboBox10.AddItem "CWH NC"
        .ComboBox10.AddItem "LVG MAINT"
        .ComboBox10.AddItem "LVG TOOL"
        .ComboBox10.AddItem "LVG FG"
        .ComboBox10.AddItem "LVG WHSE"
        .ComboBox10.AddItem "LVG DOCK"
        .ComboBox10.AddItem "OUTSOURCE"
        .ComboBox10.AddItem "SHIP FG"
        .ComboBox10.AddItem "U05 ACm "
        .ComboBox10.AddItem "U05 NC"
        .ComboBox10.AddItem "U05 RFG"
        .ComboBox10.AddItem "U05 RMm "
        .ComboBox10.AddItem "U05 WIP"
        .ComboBox10.AddItem "PACKAGING"
        .ComboBox10.AddItem "LVG PRJENG"
        .ComboBox10.AddItem "MEX MAINT"
        .ComboBox10.AddItem "MEX TOOL"
        .ComboBox10.AddItem "U21 Rm "
        .ComboBox10.AddItem "MEX FG"
        .ComboBox10.AddItem "MEX NC"
        .ComboBox10.AddItem "MEX PKG"
        .ComboBox10.AddItem "MEX REC"
        .ComboBox10.AddItem "SIMA"
        .ComboBox10.AddItem "U21 WIP"
        .ComboBox10.AddItem "U21 STAGE"
        .ComboBox10.AddItem "U17 WIP"
        .ComboBox10.AddItem "U17 COMP"
        .ComboBox10.AddItem "TACWHSE"
        .ComboBox10.AddItem "TACPRD"
        .ComboBox10.AddItem "TEMP REC"
        .ComboBox10.AddItem "SLB NONCON"
        .ComboBox10.AddItem "SLB PROD"
        .ComboBox10.AddItem "SLB FG"
        .ComboBox10.AddItem "SLB WHSE"
        .ComboBox10.AddItem "SLB SERVIC"
        .ComboBox10.AddItem "SLB RESIN"
        .ComboBox10.AddItem "SLB PACK"
        .ComboBox10.AddItem "SLB SORT"
        .ComboBox10.AddItem "SLB Rm "
        .ComboBox10.AddItem "SES"
        .ComboBox10.AddItem "SLB SRVCPK"
        .ComboBox10.AddItem "SLB TOOLS"
        .ComboBox10.AddItem "SLB PE"
        .ComboBox10.AddItem "SES CORP"
        .ComboBox10.AddItem "MEXEY"
        .ComboBox10.AddItem "SLB PCK"
        .ComboBox10.AddItem "SLB SCRAP"
        .ComboBox10.AddItem "SES PT1"
        .ComboBox10.AddItem "SLB REC 1"
        .ComboBox10.AddItem "SLB REC 2"
        .ComboBox10.AddItem "SLBLOGLOST"
        
        'mandatory field
        .ComboBox1.Clear
        .ComboBox1.AddItem "Shoot & Ship"
        .ComboBox1.AddItem "Molded Component"
        .ComboBox1.AddItem "Assembly"
        .ComboBox1.AddItem "Sub Assembly"
        
        'mandatory field
        .ComboBox14.Clear
        .ComboBox14.AddItem "U01"
        .ComboBox14.AddItem "U05"
        .ComboBox14.AddItem "U12M"
        .ComboBox14.AddItem "U12"
        .ComboBox14.AddItem "U17"
        .ComboBox14.AddItem "U18"
        .ComboBox14.AddItem "U19"
        .ComboBox14.AddItem "U20"
        .ComboBox14.AddItem "U21"
        .ComboBox14.AddItem "U40"
        .ComboBox14.AddItem "U41"
        .ComboBox14.AddItem "U42"
        .ComboBox14.AddItem "U43"
        .ComboBox14.AddItem "U55"
        
        
        'mandatory field
        .ComboBox3.Clear
        .ComboBox3.AddItem "N/A"
        .ComboBox3.AddItem "0"
        .ComboBox3.AddItem "1"
        .ComboBox3.AddItem "2"
        .ComboBox3.AddItem "3"
        .ComboBox3.AddItem "4"
        .ComboBox3.AddItem "5"
        
        'mandatory field
        .ComboBox2.Clear
        .ComboBox2.AddItem "N/A"
        .ComboBox2.AddItem "0"
        .ComboBox2.AddItem "1"
        .ComboBox2.AddItem "2"
        .ComboBox2.AddItem "3"
        .ComboBox2.AddItem "4"
        .ComboBox2.AddItem "5"
        
        'mandatory field
        .ComboBox4.Clear
        .ComboBox4.AddItem "Manual Labor"
        .ComboBox4.AddItem "Automatic"
        .ComboBox4.AddItem "Semi-Automatic"
        .ComboBox4.AddItem "Aligner Jig"
        .ComboBox4.AddItem "Assembly Labor"
        .ComboBox4.AddItem "Grommet & Seal"
        .ComboBox4.AddItem "Pin & Grommet"
        .ComboBox4.AddItem "Insert Molding"
        .ComboBox4.AddItem "No Assembly required"

        
    
    End With

End Sub

Sub ResetMoldForm()
' this load the molded section for add item on kickoff sheet

    With MoldForm
    .ComboBox2.Clear
    Dim prow As Long
    prow = Cells(Rows.count, 2).End(xlUp).Row
    Dim myTablePrs As ListObject
    Dim myArrayPrs As Variant
    Dim c As Range
    Dim pressList As String
    pressList = "PRS" & Cells(prow, 23).Value
    Set myTablePrs = Worksheets("TABLES").ListObjects(pressList)
    'myArrayPrs = myTablePrs.DataBodyRange
    For Each c In myTablePrs.DataBodyRange
        If c.Value <> vbNullString Then .ComboBox2.AddItem c.Value
    Next c
    'ComboBox2.List = myArrayPrs
    
    
    
    .ComboBox1.Clear
    .ComboBox1.AddItem "N/A"
    .ComboBox1.AddItem "0"
    .ComboBox1.AddItem "1"
    .ComboBox1.AddItem "2"
    .ComboBox1.AddItem "3"
    .ComboBox1.AddItem "4"
    .ComboBox1.AddItem "5"
    
    .ComboBox3.Clear
    .ComboBox3.AddItem "0"
    .ComboBox3.AddItem "1"
    .ComboBox3.AddItem "2"
    .ComboBox3.AddItem "3"
    .ComboBox3.AddItem "4"
    
    .ComboBox4.Clear
    .ComboBox4.AddItem "0"
    .ComboBox4.AddItem "1"
    .ComboBox4.AddItem "2"
    .ComboBox4.AddItem "3"
    .ComboBox4.AddItem "4"
    
    End With
    
    
End Sub
Sub ResetCompForm()

'this sub load the comp section of add item form

    With CompForm

'populate combo list values
Dim myArray As Variant
myArray = Worksheets("KickOFF").Range("d7:d99")

    .ComboBox1.Clear
    .ComboBox1.List = myArray
    
    .ComboBox2.Clear
    .ComboBox2.List = myArray
    
    .ComboBox3.Clear
    .ComboBox3.List = myArray
    
    .ComboBox4.Clear
    .ComboBox4.List = myArray
    
    .ComboBox5.Clear
    .ComboBox5.List = myArray
    
    .ComboBox6.Clear
    .ComboBox6.List = myArray
    
    .ComboBox7.Clear
    .ComboBox7.List = myArray
    
    .ComboBox8.Clear
    .ComboBox8.List = myArray
    
    .ComboBox9.Clear
    .ComboBox9.List = myArray
    
    .ComboBox10.Clear
    .ComboBox10.List = myArray
    


     End With

End Sub

Attribute VB_Name = "BomFormCode"
Public Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As LongPtr, ByVal y As LongPtr) As LongPtr
Public Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwmilliseconds As LongPtr)

Public Declare PtrSafe Sub mouse_event Lib "user32" (ByVal dwFlags As LongPtr, ByVal dx As Long, ByVal dy As LongPtr, ByVal cButtons As LongPtr, ByVal swextrainfo As LongPtr) '
Public Const mouseeventf_Leftdown = &H2
Public Const mouseeventf_Leftup = &H4
Public Const mouseeventf_Rightdown As Long = &H8
Public Const mouseeventf_rightup As Long = &H10

Public Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
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
Sub ClearClipboard()
    OpenClipboard (0&)
    EmptyClipboard
    CloseClipboard
End Sub

Sub setvars()
'this sub sets assigns the values of item variables which are used by other subs
t = findCell.Column

ItemNum = Cells(findCell.Row, 2).Value
orgCode = Cells(findCell.Row, 30).Value
iRow = findCell.Row
PPH = Cells(findCell.Row, 9).Value
inspecLvl = Cells(findCell.Row, 7).Value
mspakLvl = Cells(findCell.Row, 8).Value
gateLvl = Cells(findCell.Row, 31).Value
AnnealLvl = Cells(findCell.Row, 32).Value
Planner = Cells(findCell.Row, 35).Value

If orgCode = "SLB" Then BringToFront
If orgCode = "SLB" Then MsgBox "Please use SLB Cost update version for SLB ORG items"
If orgCode = "SLB" Then End

pressCode = ""
rateCode = ""

If Cells(findCell.Row, 30).Value = "MEX" Or Cells(findCell.Row, 36).Value = "MEX" Then pressCode = "m"
If Cells(findCell.Row, 30).Value = "MEX" Or Cells(findCell.Row, 36).Value = "MEX" Then rateCode = "m"
pressCode = pressCode & Left(Cells(findCell.Row, 10).Value, 1)
If Cells(findCell.Row, 19).Value = "Yes" Then pressCode = pressCode & "f"
pressCode = pressCode & "pj"
rateCode = rateCode & "pj"


pressSize = Cells(findCell.Row, 16).Value

End Sub
Sub GrabToolCost()

'this sub grabs the acct code from the LocArray sheet, the accounting code based on item type and org

Dim rng1 As Range
Dim rng2 As Range
Dim findCellRow As Range
Dim OrgCol As Range
Set rng1 = ThisWorkbook.Worksheets("LocArray").Range("i17:i21")
Set rng2 = ThisWorkbook.Worksheets("LocArray").Range("j16:n16")

Set findCellRow = rng1.Find(what:=Cells(findCell.Row, 17).Value, MatchCase:=False)

Set OrgCol = rng2.Find(what:=orgCode, MatchCase:=False)

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
Sub MainSeq()


Dim Cl As Range
Dim wrkRng As Range
Set wrkRng = ThisWorkbook.Worksheets("Kickoff Boms").Range("B9:B113")



Application.FindFormat.Clear
Application.FindFormat.Interior.ColorIndex = 2

Dim rng As Range
Set rng = ThisWorkbook.Worksheets("Kickoff Boms").Range("B9:B113")
Set findCell = rng.Find(what:="*", searchFormat:=True)
If (findCell Is Nothing) Then
    slow1
    DoEvents
    BringToFront
    MsgBox "No Unprocessed Items"
    Exit Sub
    

End If

If Not Cells(findCell.Row, 39).Value = "" Then GoTo ReStart

MasterItemCheck
RoutingLoad
RunBOM
entercost

GoTo EndLine
ReStart:
ReStartSeq

EndLine:
DoEvents
End Sub
Sub ReStartSeq()

If Cells(findCell.Row, 39).Value = "entercost" Then
    
    entercost
End If
If Cells(findCell.Row, 39).Value = "RunBOM" Then
    
    RunBOM
    entercost
End If
On Error Resume Next
If Cells(findCell.Row, 39).Value = "RoutingLoad" Then
    RoutingLoad
    RunBOM
    entercost
End If
On Error GoTo 0
slow1
DoEvents
End Sub


Sub codeGrabber()
'this sub grabs rates from tables on the Locarray sheet, based on org and lvl

Dim rng1 As Range
Set rng1 = ThisWorkbook.Worksheets("LocArray").Range("a34:a44")
Dim PressRow As Range
Set PressRow = rng1.Find(what:=pressSize, MatchCase:=False)

Dim rng2 As Range
Set rng2 = ThisWorkbook.Worksheets("LocArray").Range("a32:AJ32")
Dim pressCol As Range
Set pressCol = rng2.Find(what:=pressCode, MatchCase:=False)

Dim rng3 As Range
Set rng3 = ThisWorkbook.Worksheets("LocArray").Range("a48:h48")
Dim rateCol As Range
Set rateCol = rng3.Find(what:=rateCode, MatchCase:=False)

Dim rng4 As Range
Set rng4 = ThisWorkbook.Worksheets("LocArray").Range("a49:a54")
Dim inspecRow As Range
Set inspecRow = rng4.Find(what:=inspecLvl, MatchCase:=False)

Dim rng5 As Range
Set rng5 = ThisWorkbook.Worksheets("LocArray").Range("a58:a63")
Dim mspakRow As Range
Set mspakRow = rng5.Find(what:=mspakLvl, MatchCase:=False)

Dim rng6 As Range
Set rng6 = ThisWorkbook.Worksheets("LocArray").Range("a67:a71")
Dim gateRow As Range
Set gateRow = rng6.Find(what:=gateLvl, MatchCase:=False)

Dim rng7 As Range
Set rng7 = ThisWorkbook.Worksheets("LocArray").Range("a76:a80")
Dim annealRow As Range
Set annealRow = rng7.Find(what:=AnnealLvl, MatchCase:=False)


DoEvents


pressRate = Worksheets("LocArray").Cells(PressRow.Row, pressCol.Column).Value
inspecCode = Worksheets("LocArray").Cells(inspecRow.Row, rateCol.Column).Value
mspakCode = Worksheets("LocArray").Cells(mspakRow.Row, rateCol.Column).Value
gateCode = Worksheets("LocArray").Cells(gateRow.Row, rateCol.Column).Value
AnnealCode = Worksheets("LocArray").Cells(annealRow.Row, rateCol.Column).Value


End Sub

Sub RoutingLoad()

'this sub runs a loop through the "Kickoff Boms" sheets item list, enterting values in master item, routing, Bom, and cost elements


Dim Cl As Range
Dim wrkRng As Range
Set wrkRng = ThisWorkbook.Worksheets("Kickoff Boms").Range("B9:B113")
'For Each Cl In wrkRng


Application.FindFormat.Clear
Application.FindFormat.Interior.ColorIndex = 2

Dim rng As Range
Set rng = ThisWorkbook.Worksheets("Kickoff Boms").Range("B9:B113")
Set findCell = rng.Find(what:="*", searchFormat:=True)
If (findCell Is Nothing) Then
    BringToFront
    MsgBox "No Unprocessed Items"
    Exit Sub
    

End If

'' trying to figure out how to compare date values so I can spot check for BOM window alignment


ClickOnCornerWindow
setvars
If Not Cells(findCell.Row, 40) = "" Then MarkCode

'open Routings
slow2
Application.SendKeys ("%fw"), True

slow1
Application.SendKeys ("nb"), True
slow1


changeBM

slow1
Application.SendKeys ("1"), True

'enter item and check 00 to ensure we don't encouter item alread exsist error
slow1
Application.SendKeys ItemNum

'load routing details
If Not Cells(findCell.Row, 30).Value = "SLB" Then
    If Not Cells(findCell.Row, 30).Value = "LVG" Then
        If Not Cells(findCell.Row, 6).Value = "" Then
            Application.SendKeys "%d"
            slow1
            Application.SendKeys Cells(findCell.Row, 6).Value
            Application.SendKeys ("{Tab}")
            Application.SendKeys Cells(findCell.Row, 5).Value
            Application.SendKeys "%w2"
        End If
    End If
End If
Application.SendKeys "%fv"
slow2
Application.SendKeys ("{up}")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")

slow2
CopyCompareCell

If Not txt1 = "00" Then BringToFront
If Not txt1 = "00" Then MsgBox "out of alignment"
If Not txt1 = "00" Then End
DoEvents



slow1
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
slow1
Application.SendKeys "10", True
DoEvents
slow1
Application.SendKeys ("{Tab}"), True
Application.SendKeys " "
slow2
Application.SendKeys ("{Tab}"), True
Application.SendKeys ("{right}")
slow2
Application.SendKeys ("{Tab}"), True
Application.SendKeys ("{right}")
slow1
Application.SendKeys "Rates"

'rates section
codeGrabber
Dim rateSeq As Integer

Application.SendKeys "%r"

'If Not Cells(findCell.Row, 16) = Empty Then
slow1
rateSeq = rateSeq + 10
Application.SendKeys rateSeq
Application.SendKeys ("{Tab}")
Application.SendKeys pressRate
slow1
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
slow1
Application.SendKeys PPH

'End If
'if checks for inspec, mspak, gate, and annealing

If Cells(findCell.Row, 7).Value > 0 Then
        Application.SendKeys ("{down}") 'down arrow
        slow1
        rateSeq = rateSeq + 10
        Application.SendKeys rateSeq
        Application.SendKeys ("{Tab}")
        slow1
        Application.SendKeys inspecCode
        Application.SendKeys ("{Tab}")
        Application.SendKeys ("{Tab}")
        Application.SendKeys ("{Tab}")
        slow1
        If inspecLvl = 5 Then
            Application.SendKeys ("{Tab}")
            Application.SendKeys PPH
        Else
            Application.SendKeys "1"
            Application.SendKeys ("{Tab}")
            Application.SendKeys "1"
        End If
End If

If Cells(findCell.Row, 8).Value > 0 Then
        Application.SendKeys ("{down}") 'down arrow
        slow1
        rateSeq = rateSeq + 10
        Application.SendKeys rateSeq
        Application.SendKeys ("{Tab}")
        slow1
        Application.SendKeys mspakCode
        Application.SendKeys ("{Tab}")
        Application.SendKeys ("{Tab}")
        Application.SendKeys ("{Tab}")
        slow1
        If mspakLvl = 5 Then
            Application.SendKeys ("{Tab}")
            Application.SendKeys PPH
        Else
            Application.SendKeys "1"
            Application.SendKeys ("{Tab}")
            Application.SendKeys "1"
        End If
End If

If Cells(findCell.Row, 31).Value > 0 Then
        Application.SendKeys ("{down}") 'down arrow
        slow1
        rateSeq = rateSeq + 10
        Application.SendKeys rateSeq
        Application.SendKeys ("{Tab}")
        slow1
        Application.SendKeys gateCode
        Application.SendKeys ("{Tab}")
        Application.SendKeys ("{Tab}")
        Application.SendKeys ("{Tab}")
        slow1
        If gateLvl = 5 Then
            Application.SendKeys ("{Tab}")
            Application.SendKeys PPH
        Else
            Application.SendKeys "1"
            Application.SendKeys ("{Tab}")
            Application.SendKeys "1"
        End If
End If
If Cells(findCell.Row, 32).Value > 0 Then
        Application.SendKeys ("{down}") 'down arrow
        slow1
        rateSeq = rateSeq + 10
        Application.SendKeys rateSeq
        Application.SendKeys ("{Tab}")
        slow1
        Application.SendKeys AnnealCode
        Application.SendKeys ("{Tab}")
        Application.SendKeys ("{Tab}")
        Application.SendKeys ("{Tab}")
        slow1
        If AnnealLvl = 5 Then
            Application.SendKeys ("{Tab}")
            Application.SendKeys PPH
        Else
            Application.SendKeys "1"
            Application.SendKeys ("{Tab}")
            Application.SendKeys "1"
        End If
End If

'close OR box
    'switch to window 2
Application.SendKeys "%fs"
Application.SendKeys "%w2"
'load shld res
slow1
ClickmainRoutBox
Dim shldSeq As Integer
If Not Cells(findCell.Row, 3).Value = "" Then
    
    Application.SendKeys ("{down}") 'down arrow
    slow1
    Application.SendKeys "%n"
    'slow1
    'Application.SendKeys "20"
    'DoEvents
    slow1
    Application.SendKeys ("{Tab}"), True
    Application.SendKeys " "
    slow2
    Application.SendKeys ("{Tab}"), True
    Application.SendKeys ("{right}")
    slow2
    Application.SendKeys ("{Tab}"), True
    Application.SendKeys ("{right}")
    slow1
    Application.SendKeys "SHL"
    Application.SendKeys ("{Tab}"), True
    slow1
    Application.SendKeys "%r"
    slow1
    Application.SendKeys "10"
    Application.SendKeys ("{Tab}"), True
    Application.SendKeys Cells(findCell.Row, 3).Value
    slow1
    Application.SendKeys ("{Tab}"), True
    Application.SendKeys ("{Tab}")
    Application.SendKeys ("{Tab}")
    Application.SendKeys ("{Tab}")
    'Application.SendKeys ("{Tab}")
    slow1
    Application.SendKeys PPH
    slow1
    Application.SendKeys ("{Tab}"), True
    Application.SendKeys ("{Tab}")
    Application.SendKeys ("{Tab}")
    Application.SendKeys ("{Tab}")
    slow1
    Application.SendKeys ("{Tab}")
    slow1
    Application.SendKeys "yes"
    Application.SendKeys ("{Tab}")
    'Stop
'End If
    'press setup,  could fail if can't find
If Cells(findCell.Row, 30).Value = "CNL" Then
    If Not Cells(findCell.Row, 33).Value = "" Then
        Application.SendKeys ("{down}") 'down arrow
        slow1
        Application.SendKeys "20"
         Application.SendKeys ("{Tab}"), True
        Application.SendKeys "Press"
        Application.SendKeys " "
        Application.SendKeys Cells(findCell.Row, 33).Value
        slow1
        Application.SendKeys ("{Tab}")
        Application.SendKeys ("{Tab}")
        Application.SendKeys ("{Tab}")
        Application.SendKeys ("{Tab}")
        slow1
        Application.SendKeys PPH
        slow1
        Application.SendKeys ("{Tab}"), True
        Application.SendKeys ("{Tab}")
        Application.SendKeys ("{Tab}")
        Application.SendKeys ("{Tab}")
        slow1
        Application.SendKeys ("{Tab}")
        slow1
        Application.SendKeys "yes"
        Application.SendKeys ("{Tab}")
    End If
End If
If Cells(findCell.Row, 30).Value = "LVG" Then
    If Not Cells(findCell.Row, 33).Value = "" Then
        Application.SendKeys ("{down}") 'down arrow
        slow1
        Application.SendKeys "20"
         Application.SendKeys ("{Tab}"), True
        Application.SendKeys "Press"
        Application.SendKeys " "
        Application.SendKeys Cells(findCell.Row, 33).Value
        slow1
        Application.SendKeys ("{Tab}")
        Application.SendKeys ("{Tab}")
        Application.SendKeys ("{Tab}")
        Application.SendKeys ("{Tab}")
        slow1
        Application.SendKeys PPH
        slow1
        Application.SendKeys ("{Tab}"), True
        Application.SendKeys ("{Tab}")
        Application.SendKeys ("{Tab}")
        Application.SendKeys ("{Tab}")
        slow1
        Application.SendKeys ("{Tab}")
        slow1
        Application.SendKeys "yes"
        Application.SendKeys ("{Tab}")
    End If
End If
Application.SendKeys "%w2"
End If

Application.SendKeys "%fs"
slow2
Application.SendKeys ("~")
slow2
Application.SendKeys "%fs"
'CloseRout
slow2
Application.SendKeys "%fc"
slow2

Cells(findCell.Row, 39).Value = "RunBOM"
'Next Cl



SkipToEnd:
'End
End Sub

Sub RunBOM()
'this sub is to navigate through the BOM form and enter components, resin and or packaging.

Set rng = ThisWorkbook.Worksheets("Kickoff Boms").Range("B9:B113")
Application.FindFormat.Clear
Application.FindFormat.Interior.ColorIndex = 2
Set findCell = rng.Find(what:="*", searchFormat:=True)
If (findCell Is Nothing) Then
    BringToFront
    MsgBox "No Unprocessed Items"
    Exit Sub
End If
slow1
ClickOnCornerWindow

setvars

Application.SendKeys "%fw"
slow1
Application.SendKeys "nb"
slow2
changeBM


Application.SendKeys "2"

slow1
Application.SendKeys ItemNum
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
On Error GoTo 0
txt3 = DateValue(WorksheetFunction.Text(txt4, "mm/dd/yyyy"))
If Not txt3 = txt2 Then Stop
If Not txt3 = txt2 Then EndBOM

ClearClipboard

' Resin Section

Dim itemseq As Integer

Dim resinArray As Variant
Dim weightArray As Variant

resinArray = Split(Cells(findCell.Row, 20).Value, Chr(10))
weightArray = Split(Cells(findCell.Row, 21).Value, Chr(10))

slow1
Application.SendKeys "+{PGDN}"

slow1


Dim I1 As Integer
Dim I2 As Integer
Dim I3 As Integer
Dim L1 As Integer
Dim L2 As Integer
Dim L3 As Integer
L1 = UBound(resinArray) + 1

For I1 = 1 To L1
If Not resinArray(I1 - 1) = "" Then
    
        itemseq = itemseq + 10
        Application.SendKeys itemseq
        Application.SendKeys ("{Tab}")
        Application.SendKeys "10"
        Application.SendKeys ("{Tab}")
        Application.SendKeys resinArray(I1 - 1)
        slow1
        Application.SendKeys ("{Tab}")
        Application.SendKeys ("{Tab}")
        ClearClipboard
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
        
        Application.SendKeys weightArray(I1 - 1)
        slow1
        Application.SendKeys ("{Tab}")
        Application.SendKeys ("{Tab}")
        Application.SendKeys ("{Tab}")
        Application.SendKeys ("{right}")
        slow2
        Application.SendKeys ("{Tab}")
        Application.SendKeys ("{Tab}")
        Application.SendKeys ("{right}")
        'slow2
        'Application.SendKeys ("{Tab}")
        'Application.SendKeys ("{right}")
        'Application.SendKeys ("{Tab}")
        'slow2
        'If itemseq > 10 Then Application.SendKeys ("{Tab}")
       
        For r = 1 To 20
            slow1
            CopyCompareCell
            On Error Resume Next
            If txt1 = 100 Then r = 21
            On Error GoTo 0
            
            Application.SendKeys ("{Tab}")
            
            slow1
        Next
        Application.SendKeys ("{Tab}")
        Application.SendKeys ".99"
        
        Application.SendKeys ("{Tab}")
        Application.SendKeys ("{Tab}")
        Application.SendKeys ("{Tab}")
        slow1
        
        If Cells(findCell.Row, 30).Value = "CNL" Then
            If i = 1 Then
               CopyCompareCell
                If txt1 = "Assembly Pull" Then Application.SendKeys "Push"
            End If
        End If
        If Cells(findCell.Row, 30).Value = "SLB" Then
            If i = 2 Then
               CopyCompareCell
                If txt1 = "Assembly Pull" Then Application.SendKeys "Bulk"
            End If
        End If
        
        CopyCompareCell
        If Cells(findCell.Row, 30).Value = "SLB" Then
            If txt1 = "Assembly Pull" Then
                Application.SendKeys ("{Tab}")
                Application.SendKeys "SLB PROD"
            End If
        End If
        If Cells(findCell.Row, 30).Value = "LVG" Then
            If txt1 = "Assembly Pull" Then
                Application.SendKeys ("{Tab}")
                Application.SendKeys "U05 RMM"
                Application.SendKeys ("{Tab}")
                Application.SendKeys "WM05.01.00"
            End If
        End If
Stop
slow1
Application.SendKeys ("{down}")
End If
Next

DoEvents
slow1
'Componenets
Dim compArray As Variant
Dim useCArray As Variant
Dim complocArray As Variant
Dim compSubInvArray As Variant


compArray = Split(Cells(findCell.Row, 26).Value, Chr(10))
useCArray = Split(Cells(findCell.Row, 27).Value, Chr(10))
complocArray = Split(Cells(findCell.Row, 28).Value, Chr(10))
compSubInvArray = Split(Cells(findCell.Row, 29).Value, Chr(10))

L2 = UBound(compArray) + 1

For I2 = 1 To L2
    If I2 > L2 Then GoTo SkipComp
    itemseq = itemseq + 10
    Application.SendKeys itemseq
    Application.SendKeys ("{Tab}")
    Application.SendKeys "10"
    Application.SendKeys ("{Tab}")
    Application.SendKeys compArray(I2 - 1)
    slow1
    Application.SendKeys ("{Tab}")
    ClearClipboard
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
    Application.SendKeys useCArray(I2 - 1)
   
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
    If Cells(findCell.Row, 30) = "SLB" Then
        If Cells(findCell.Row, 34) = True Then
            Application.SendKeys ("{Tab}")
            If I2 = 1 Then Application.SendKeys "Phantom"
        Else
        Application.SendKeys ("{Tab}")
        Application.SendKeys ("{Tab}")
        Application.SendKeys "SLB PROD"
        End If
        
    End If
slow1
Application.SendKeys ("{Tab}")
'If I2 > 1 Then Application.SendKeys ("{Tab}")
Dim Val1 As Integer
Val1 = UBound(compSubInvArray) + 1
If Not I2 > Val1 Then
    slow2
    Application.SendKeys compSubInvArray(I2 - 1)
    slow1
    
End If
slow1
Application.SendKeys ("+{Tab}")
Application.SendKeys ("{Tab}")
slow1
DoEvents
slow2
CopyCompareCell
On Error Resume Next
If txt1 = compSubInvArray(I2 - 1) Then Application.SendKeys ("{Tab}")
On Error GoTo 0
slow2
If Not I2 > Val1 Then
    slow1
    'Application.SendKeys ("{Tab}")
    Application.SendKeys complocArray(I2 - 1)
End If
slow2
Application.SendKeys ("{Tab}")


'End If

slow1
Application.SendKeys ("{down}")



Next
slow1
Application.SendKeys ("{up}")
slow1
If itemseq > 10 Then
    CopyCompareCell
    If Not txt1 = itemseq Then EndBOM
End If
Application.SendKeys ("{down}")
slow1

SkipComp:

'load packaging

Dim packArray As Variant
Dim usePArray As Variant

packArray = Split(Cells(findCell.Row, 24).Value, Chr(10))
usePArray = Split(Cells(findCell.Row, 25).Value, Chr(10))

L3 = UBound(packArray) + 1

For I3 = 1 To L3

    itemseq = itemseq + 10
    Application.SendKeys itemseq
    slow1
    Application.SendKeys ("{Tab}")
    Application.SendKeys "10"
    Application.SendKeys ("{Tab}")
    slow1
    Application.SendKeys packArray(I3 - 1)
    slow1
    Application.SendKeys ("{Tab}")
    ClearClipboard
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
    Application.SendKeys usePArray(I3 - 1)
    
    
    slow1
    'Application.SendKeys ("{Tab}")
    slow1
       If Left(packArray(I3 - 1), 1) = "R" Then
       
        Application.SendKeys ("{Tab}")
        Application.SendKeys ("{Tab}")
        Application.SendKeys ("{Tab}")
        slow1
        Application.SendKeys ("{Tab}")
        Application.SendKeys ("{Tab}")
        Application.SendKeys ("{Tab}")
        slow1
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
        slow1
        Application.SendKeys "BULK"
        End If
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
       
    End If
    slow1
Application.SendKeys ("{down}")

txt1 = 0
Next
Application.SendKeys ("{up}")
CopyCompareCell
If Not txt1 = itemseq Then EndBOM
slow1

Application.SendKeys "%fs"
slow2
Application.SendKeys "%fc"

Cells(findCell.Row, 39).Value = "entercost"

End Sub
Sub MasterItemCheck()
'this sub is for navigating through the master item form, making the item active, checking various feilds and assigning to orgs and planner.

Dim Cl As Range
Dim wrkRng As Range
Set wrkRng = ThisWorkbook.Worksheets("Kickoff Boms").Range("B9:B113")



Application.FindFormat.Clear
Application.FindFormat.Interior.ColorIndex = 2

Dim rng As Range
Set rng = ThisWorkbook.Worksheets("Kickoff Boms").Range("B9:B113")
Set findCell = rng.Find(what:="*", searchFormat:=True)
If (findCell Is Nothing) Then
    BringToFront
    MsgBox "No Unprocessed Items"
    Exit Sub
    

End If

slow2

ClickOnCornerWindow

setvars
codeGrabber
Application.SendKeys "%fw"
slow2
Application.SendKeys "ni"
slow2
Application.SendKeys "%tt"
Application.SendKeys "1"
slow2
Application.SendKeys "CC"
slow2
Application.SendKeys "%vf"
slow2
Application.SendKeys ItemNum
Application.SendKeys "%i"
slow2
Application.SendKeys "%tf"
slow2
Application.SendKeys "item s"
slow2
Application.SendKeys "Active"
slow1
Application.SendKeys ("{Tab}")
slow1
Application.SendKeys "%tf"
slow2
Application.SendKeys "checkm"
slow1
Application.SendKeys ("{F5}")
slow1
Application.SendKeys " "
slow2

Application.SendKeys "%tf"
slow1
Application.SendKeys "item s"
slow2
CopyCompareCell
If Not txt1 = "Active" Then BringToFront
If Not txt1 = "Active" Then MsgBox "out of alignment, check if item is setup"
slow2
If Not txt1 = "Active" Then End
Application.SendKeys "%fs"
slow2

Application.SendKeys "%to"
If orgCode = "CNL" Then
    Application.SendKeys ("{down}")
    Application.SendKeys "%o"
    slow2
End If
If orgCode = "GWH" Then
    Application.SendKeys ("{down}")
    Application.SendKeys ("{down}")
    Application.SendKeys ("{down}")
    Application.SendKeys "%o"
    slow2
End If
If orgCode = "LVG" Then
    Application.SendKeys ("{down}")
    Application.SendKeys ("{down}")
    Application.SendKeys ("{down}")
    Application.SendKeys ("{down}")
    Application.SendKeys "%o"
    slow2
End If
If orgCode = "MEX" Then
    Application.SendKeys ("{down}")
    Application.SendKeys ("{down}")
    Application.SendKeys ("{down}")
    Application.SendKeys ("{down}")
    Application.SendKeys ("{down}")
    Application.SendKeys "%o"
    slow2
End If
If orgCode = "SLB" Then
    Application.SendKeys ("{down}")
    Application.SendKeys ("{down}")
    Application.SendKeys ("{down}")
    Application.SendKeys ("{down}")
    Application.SendKeys ("{down}")
    Application.SendKeys ("{down}")
    Application.SendKeys ("{down}")
    Application.SendKeys "%o"
    slow2
End If
slow2
slow2
slow2
Application.SendKeys "%tf"
slow2
Application.SendKeys "Planner"
slow2
Application.SendKeys Planner
slow2
Application.SendKeys "%fs"
Application.SendKeys ("{Tab}")
slow2
'Application.SendKeys ("{F4}")
Application.SendKeys "%fc"
slow1
'Application.SendKeys ("{F4}")
Application.SendKeys "%fc"
slow2
slow1
If orgCode = "CNL" Then ADDGWHPlanner

Cells(findCell.Row, 39).Value = "RoutingLoad"

End Sub
Sub MarkCode()



Application.SendKeys "%to"
Application.SendKeys "%tu"
slow1
Application.SendKeys "i"
Application.SendKeys "ch"
slow1
Application.SendKeys "cnl"
Application.SendKeys "%tc"
slow1
Application.SendKeys "%to"
Application.SendKeys "%tu"
slow1
Application.SendKeys "i"
slow1
Application.SendKeys "i"
Application.SendKeys "c"
slow1
Application.SendKeys ("~")
slow1

Application.SendKeys "%vf"
slow1
Application.SendKeys "hondam"
Application.SendKeys "%a"
slow1

Application.SendKeys "%fn"
Application.SendKeys ItemNum
slow1
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
slow1
Application.SendKeys Cells(findCell.Row, 40).Value
slow1
Application.SendKeys "%fs"
slow1
Application.SendKeys "%fc"
slow1

Stop
End Sub
Sub ADDGWHPlanner()

Application.SendKeys "1"
slow2
slow2
Application.SendKeys "%vf"
slow2
Application.SendKeys ItemNum
Application.SendKeys "%i"
slow2
Application.SendKeys "%tf"
slow2
Application.SendKeys "item s"
slow2
CopyCompareCell
If Not txt1 = "Active" Then BringToFront
If Not txt1 = "Active" Then MsgBox "out of alignment, check if item is setup"
slow2
If Not txt1 = "Active" Then End
Application.SendKeys "%to"
slow2
Application.SendKeys ("{down}")
Application.SendKeys ("{down}")
Application.SendKeys ("{down}")
Application.SendKeys "%o"
slow2
slow2
slow2
Application.SendKeys "%tf"
slow2
Application.SendKeys "Planner"
slow2
If Left(Planner, 5) = "DAZAR" Then
    Application.SendKeys "DAZAR-GWH"
    Else
     Application.SendKeys Planner
     Application.SendKeys ("{Tab}")
End If

slow2
Application.SendKeys "%tf"
Application.SendKeys "Planner"
slow1
Application.SendKeys "%f"
slow1
CopyCompareCell

If Left(Planner, 5) = "DAZAR" Then GoTo skipDAZ
If Not txt1 = Planner Then BringToFront
If Not txt1 = Planner Then MsgBox "Error in Adding Planner to GWH, Ending Script"
If Not txt1 = Planner Then Stop
slow1
skipDAZ:
Application.SendKeys "%fs"
Application.SendKeys ("{Tab}")
slow2
slow1
Application.SendKeys ("{F4}")
slow2
slow1
Application.SendKeys "%fc"
slow1
End Sub

Sub CloseRout()
'this is for closing the routing window, possible I can eliminate this through closing and reopening forms
Dim oLeft As Long
oLeft = 745

Dim OTop As Long
OTop = 93

SetCursorPos oLeft, OTop

mouse_event mouseeventf_Leftdown, 0, 0, 0, 0
mouse_event mouseeventf_Leftup, 0, 0, 0, 0


End Sub
Sub TransferFinishtoAIFSheet()

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
        .Cells((iRow.Row), 5) = "Kickoff"
        .Cells((iRow.Row), 7) = "Kickoff"
        If Cells(findCell.Row, 30) = "CNL" Then
            .Cells((iRow.Row), 3) = "CNL"
            .Cells((iRow.Row), 4) = "107"
        End If
        If Cells(findCell.Row, 30) = "GWH" Then
            .Cells((iRow.Row), 3) = "GWH"
            .Cells((iRow.Row), 4) = "107"
        End If
        If Cells(findCell.Row, 30) = "LVG" Then
            .Cells((iRow.Row), 3) = "LVG"
            .Cells((iRow.Row), 4) = "105"
        End If
        If Cells(findCell.Row, 30) = "MEX" Then
            .Cells((iRow.Row), 3) = "MEX"
            .Cells((iRow.Row), 4) = "104"
        End If
        If Cells(findCell.Row, 30) = "SLB" Then
            .Cells((iRow.Row), 3) = "SLB"
            .Cells((iRow.Row), 4) = "109"
        End If
        .Cells((iRow.Row), 11) = Cells(findCell.Row, 12).Value
    End With
        


End Sub

Sub entercost()

'this sub is for entereing cost elements in the item cost form

Application.FindFormat.Clear
Application.FindFormat.Interior.ColorIndex = 2
Dim rng As Range
Set rng = ThisWorkbook.Worksheets("Kickoff Boms").Range("B9:B113")
Set findCell = rng.Find(what:="*", searchFormat:=True)
If (findCell Is Nothing) Then
    BringToFront
    MsgBox "No Unprocessed Items"
    Exit Sub
End If
''


ExitCon = False
setvars
codeGrabber
GrabToolCost

ClickOnCornerWindow
slow2
Application.SendKeys ("%f+w"), True
slow2

Application.SendKeys ("cc"), True
slow2

changeOrg
slow2
Application.SendKeys ("1")

slow2

Application.SendKeys ItemNum
slow1
Application.SendKeys ("~")
slow1

Application.SendKeys ("{Tab}")

CopyCompareCell
If Not txt1 = "Frozen" Then BringToFront
If Not txt1 = "Frozen" Then ExitCon = True
If Not txt1 = "Frozen" Then MsgBox "Out of Alignment Closing, please assign remaining Orgs manually"

DoEvents
Application.Wait (Now + TimeValue("00:00:01"))
If ExitCon = True Then GoTo SkipToEnd

Application.SendKeys ("{down}")
Application.SendKeys ItemNum
slow1
Application.SendKeys ("{Tab}")
Application.SendKeys "Kickoff"
slow1
Application.SendKeys ("%c")

'specific subs are called based on type of item
If Cells(findCell.Row, 4) = "Shoot & Ship" Then enterShootShip
If Cells(findCell.Row, 4) = "Molded Component" Then enterMoldedComp
If Cells(findCell.Row, 4) = "Assembly" Then enterAssembly
If Cells(findCell.Row, 4) = "Sub Assembly" Then enterAssembly


slow2
Application.SendKeys ("%fs")


slow1
Application.SendKeys ("{F4}")
slow2
Application.SendKeys ("%w1")
TransferFinishtoAIFSheet
findCell.Interior.ColorIndex = 4

Cells(findCell.Row, 39) = ""
Set findCell = Nothing
BringToFront
SkipToEnd:



End Sub


Sub CopyCompareCell()
'this sub activates a crtl C action and then pulls the info in the windows clipboard and stores it in a variable so it can be evaluated at various intervals
Dim ClipObj As New DataObject

ClipObj.SetText Text:=Empty
ClipObj.PutInClipboard
Application.SendKeys ("^c")
DoEvents
Application.Wait (Now + TimeValue("00:00:01"))


ClipObj.GetFromClipboard
On Error Resume Next
    txt1 = ClipObj.GetText(1)
On Error GoTo 0

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
Public Sub ClickmainRoutBox()
' added this function because routing form was not easily navigatable, will need to re-examine.
Dim oLeft As Long
oLeft = 170  'last desktop size
'oLeft = 80   'revised desktop

Dim OTop As Long
OTop = 400

SetCursorPos oLeft, OTop

mouse_event mouseeventf_Leftdown, 0, 0, 0, 0
mouse_event mouseeventf_Leftup, 0, 0, 0, 0


End Sub

Sub slow1()
'sub for adding delay in procedures, delays needed to keep step from overrunning oracle and causing misalignment.
DoEvents
Application.Wait (Now + TimeValue("00:00:01"))

End Sub
Sub slow2()
'2 sec version of former sub
DoEvents
Application.Wait (Now + TimeValue("00:00:02"))

End Sub
Sub changeOrg()
'this sub is for handling changing of orgs
orgCode = Cells(iRow, 30).Value
slow1

Application.SendKeys ("5"), True
slow1

Dim c As String
If orgCode = "CNL" Then c = "cn"
If orgCode = "GWH" Then c = "g"
If orgCode = "LVG" Then c = "l"
If orgCode = "MEX" Then c = "Me"
If orgCode = "SLB" Then c = "s"

Application.SendKeys (c), True

End Sub
Sub changeBM()
'this is for changing orgs in Bom and routing section
orgCode = Cells(iRow, 30).Value
slow1

Application.SendKeys ("3"), True
slow1

Dim c As String
If orgCode = "CNL" Then c = "cn"
If orgCode = "GWH" Then c = "g"
If orgCode = "LVG" Then c = "l"
If orgCode = "MEX" Then c = "Me"
If orgCode = "SLB" Then c = "s"

Application.SendKeys (c), True

End Sub


Sub BomFormReset()
'this sub sets and re-sets the add item form for kickoffs
'Dim iRow As Long

'iRow = [Counta(Kickoff Boms!B:B)]

    With BomForm
        
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
        .ComboBox4.AddItem "Hand"
        .ComboBox4.AddItem "Automatic"
        .ComboBox4.AddItem "Semi-Automatic"
        
        'mandatory field
        .ComboBox12.Clear
        .ComboBox12.AddItem "ALFONSO"
        .ComboBox12.AddItem "ALONDRA"
        .ComboBox12.AddItem "ANDRE D"
        .ComboBox12.AddItem "ANGELA C"
        .ComboBox12.AddItem "AUSTIN S"
        .ComboBox12.AddItem "BARNESCA"
        .ComboBox12.AddItem "BELINDA18"
        .ComboBox12.AddItem "BISHOPC"
        .ComboBox12.AddItem "BRISSA"
        .ComboBox12.AddItem "BROWNINGT"
        .ComboBox12.AddItem "CAHILL18M"
        .ComboBox12.AddItem "CARTWRIGHT"
        .ComboBox12.AddItem "CHAD D"
        .ComboBox12.AddItem "CHILDRESST"
        .ComboBox12.AddItem "CHRYS-COMP"
        .ComboBox12.AddItem "CHRYS-FG"
        .ComboBox12.AddItem "CHRYSTAL"
        .ComboBox12.AddItem "CHRYSTAL1"
        .ComboBox12.AddItem "CLAWSONW"
        .ComboBox12.AddItem "CLAWSONW1"
        .ComboBox12.AddItem "CORAL"
        .ComboBox12.AddItem "CRAIGS"
        .ComboBox12.AddItem "CRAIGS1"
        .ComboBox12.AddItem "CRISTIAN"
        .ComboBox12.AddItem "CRYSTAL"
        .ComboBox12.AddItem "CRYSTAL OS"
        .ComboBox12.AddItem "CZAR"
        .ComboBox12.AddItem "DAVE R"
        .ComboBox12.AddItem "DAVIDJ"
        .ComboBox12.AddItem "DAVIDSONN"
        .ComboBox12.AddItem "DAVILAJ"
        .ComboBox12.AddItem "DAZAR-CNL"
        .ComboBox12.AddItem "DAZAR-LVG"
        .ComboBox12.AddItem "DAZAR-MEX"
        .ComboBox12.AddItem "DAZAR-SLB"
        .ComboBox12.AddItem "DIEDRICK K"
        .ComboBox12.AddItem "DIEDRICK K "
        .ComboBox12.AddItem "DOUBLEA"
        .ComboBox12.AddItem "EDWARD"
        .ComboBox12.AddItem "EDWARD W"
        .ComboBox12.AddItem "ELIZABETH"
        .ComboBox12.AddItem "ERIKA C"
        .ComboBox12.AddItem "EVAN M"
        .ComboBox12.AddItem "GIOVANA"
        .ComboBox12.AddItem "HOLLY"
        .ComboBox12.AddItem "HOWELLC"
        .ComboBox12.AddItem "HOWELLC1"
        .ComboBox12.AddItem "JACK L"
        .ComboBox12.AddItem "JAVIER"
        .ComboBox12.AddItem "JAZMIN"
        .ComboBox12.AddItem "JEREMY"
        .ComboBox12.AddItem "JEREMY R"
        .ComboBox12.AddItem "JOANNA"
        .ComboBox12.AddItem "JOSE"
        .ComboBox12.AddItem "KAREN"
        .ComboBox12.AddItem "KAREN R"
        .ComboBox12.AddItem "KEVIN K"
        .ComboBox12.AddItem "LEA"
        .ComboBox12.AddItem "LILIA D"
        .ComboBox12.AddItem "LILLIA D"
        .ComboBox12.AddItem "LOGAN"
        .ComboBox12.AddItem "LUIS"
        .ComboBox12.AddItem "LUIS-ALBER"
        .ComboBox12.AddItem "LUZ P "
        .ComboBox12.AddItem "MACKEN-CNL"
        .ComboBox12.AddItem "MACKEN-GWH"
        .ComboBox12.AddItem "MAR-COMP"
        .ComboBox12.AddItem "MARCOU"
        .ComboBox12.AddItem "MARILEE"
        .ComboBox12.AddItem "MARILEE-TA"
        .ComboBox12.AddItem "MARILE-MEX"
        .ComboBox12.AddItem "MARIN B"
        .ComboBox12.AddItem "MEADON"
        .ComboBox12.AddItem "MIGUEL"
        .ComboBox12.AddItem "NCMPLANNER"
        .ComboBox12.AddItem "NO PLANNER"
        .ComboBox12.AddItem "OSBUYBACK"
        .ComboBox12.AddItem "PEGASUS"
        .ComboBox12.AddItem "PLUMMERT"
        .ComboBox12.AddItem "RAFAEL"
        .ComboBox12.AddItem "ROBYN"
        .ComboBox12.AddItem "ROBYN-COM"
        .ComboBox12.AddItem "ROBYN-COMP"
        .ComboBox12.AddItem "ROBYN-MEX"
        .ComboBox12.AddItem "ROBYN-MEX1"
        .ComboBox12.AddItem "ROBYN-OS"
        .ComboBox12.AddItem "SALAZARL"
        .ComboBox12.AddItem "SCALFC"
        .ComboBox12.AddItem "SES"
        .ComboBox12.AddItem "STAHLS"
        .ComboBox12.AddItem "STRATTONZ"
        .ComboBox12.AddItem "STUMBOK"
        .ComboBox12.AddItem "STUMBOK1"
        .ComboBox12.AddItem "SUAREZ-U18"
        .ComboBox12.AddItem "SWITZERA"
        .ComboBox12.AddItem "SUNNY"
        .ComboBox12.AddItem "TAKAKO"
        .ComboBox12.AddItem "TRACY"
        .ComboBox12.AddItem "VWBROWNING"
        .ComboBox12.AddItem "VWCECIL"
        .ComboBox12.AddItem "WAKEFIELDA"
        .ComboBox12.AddItem "WES R"
        .ComboBox12.AddItem "WISNIEWSKI"
        .ComboBox12.AddItem "YASMIN"
        .ComboBox12.AddItem "ZIMMERLEC"

        .ComboBox13.Clear
        .ComboBox13.AddItem "3M CANADA"
        .ComboBox13.AddItem "3M DO BRASIL LTDA"
        .ComboBox13.AddItem "3M DO BRASIL LTDA."
        .ComboBox13.AddItem "3M MEXICO, S.A. DE C.V."
        .ComboBox13.AddItem "3M PANAMA PACIFICO S.DE R.L"
        .ComboBox13.AddItem "6S PRODUCTS"
        .ComboBox13.AddItem "A&K FINISHING INC"
        .ComboBox13.AddItem "A. KAYSER AUTOMOTIVE SYSTEMS GMBH"
        .ComboBox13.AddItem "A.P. PLASMAN"
        .ComboBox13.AddItem "A-1 FASTENER INC"
        .ComboBox13.AddItem "AB KUNSHAN Plastic Mold CO.,LTD"
        .ComboBox13.AddItem "ABC AUTOMOTIVE SYSTEMS INC"
        .ComboBox13.AddItem "ABC CLIMATE CONTROL SYSTEMS"
        .ComboBox13.AddItem "ABC GROUP"
        .ComboBox13.AddItem "ABC GROUP PRODUCT DEVELOPMENT"
        .ComboBox13.AddItem "ABC INTERIOR SYSTEMS INC"
        .ComboBox13.AddItem "ABC TECHNOLOGIES"
        .ComboBox13.AddItem "A-BRITE LP"
        .ComboBox13.AddItem "ACCURIDE INTERNATIONAL INC."
        .ComboBox13.AddItem "ACCUDYN DE MEX"
        .ComboBox13.AddItem "ACF COMPONENTS & FASTENERS, INC."
        .ComboBox13.AddItem "ACSCO"
        .ComboBox13.AddItem "ACUMENT GLOBAL TECHNOLOGIES"
        .ComboBox13.AddItem "ADAC PLASTICS INC"
        .ComboBox13.AddItem "ADELL GROUP, INC."
        .ComboBox13.AddItem "ADIENT AUTOMOTIVE ARGENTINA S.R.L."
        .ComboBox13.AddItem "ADIENT CLANTON INC."
        .ComboBox13.AddItem "ADIENT COMPONENTS LTD. & CO. KG"
        .ComboBox13.AddItem "ADIENT DO BRASIL BANCOS AUTOMOTIVOS LTDA"
        .ComboBox13.AddItem "ADIENT ELDON INC."
        .ComboBox13.AddItem "ADIENT MEXICO AUTOMOTRIZ, S DE R.L. DE C.V."
        .ComboBox13.AddItem "ADIENT MEXICO AUTOMOTRIZ, S. DE R.L. DE C.V."
        .ComboBox13.AddItem "ADIENT QUERETARO S. DE R.L. DE C.V."
        .ComboBox13.AddItem "ADIENT SOUTH AFRICA PTY LTD"
        .ComboBox13.AddItem "ADIENT US LLC"
        .ComboBox13.AddItem "ADKEV, INC."
        .ComboBox13.AddItem "ADVANCED ASSEMBLY LLC"
        .ComboBox13.AddItem "ADVANCE COMPONENTS"
        .ComboBox13.AddItem "ADVANCE ENGINEERING"
        .ComboBox13.AddItem "ADVANCED INTERIOR SOLUTIONS"
        .ComboBox13.AddItem "ADVANCED PLASTICS, INC."
        .ComboBox13.AddItem "ADVANTAGE MANUFACTURING CORP."
        .ComboBox13.AddItem "AER MANUFACTURING L.P."
        .ComboBox13.AddItem "AERO PRODUCTS COMPANY"
        .ComboBox13.AddItem "AETHER APPAREL"
        .ComboBox13.AddItem "AFTECH"
        .ComboBox13.AddItem "AGC AUTOMOTIVE MEXICO"
        .ComboBox13.AddItem "AGM AUTOMOTIVE"
        .ComboBox13.AddItem "AGM AUTOMOTIVE LLC"
        .ComboBox13.AddItem "AGORA EDGE"
        .ComboBox13.AddItem "AGORA SALES"
        .ComboBox13.AddItem "AINAK, INC."
        .ComboBox13.AddItem "AIR DESIGN"
        .ComboBox13.AddItem "AISAN AUTOPARTES MEXICO SA DE CV"
        .ComboBox13.AddItem "AISIN AUTOMOTIVE GUANAJUATO S.A. DE C.V."
        .ComboBox13.AddItem "AISIN AUTOMOTIVE LTDA."
        .ComboBox13.AddItem "AISIN CANADA INC."
        .ComboBox13.AddItem "AISIN ELECTRONICS ILLINOIS, LLC"
        .ComboBox13.AddItem "AISIN LIGHT METALS"
        .ComboBox13.AddItem "AISIN MEXICANA S.A. DE C.V."
        .ComboBox13.AddItem "AISIN MFG. ILLINOIS"
        .ComboBox13.AddItem "AISIN NORTH CAROLINA CORPORATION"
        .ComboBox13.AddItem "AISIN TEXAS CORPORATION"
        .ComboBox13.AddItem "AKRON POLYMER PRODUCTS"
        .ComboBox13.AddItem "AKWEL MEXICO USA, INC."
        .ComboBox13.AddItem "ALEX PRODUCTS, INC."
        .ComboBox13.AddItem "ALFIELD INDUSTRIES"
        .ComboBox13.AddItem "ALFMEIER FRIEDRICHS & RATH"
        .ComboBox13.AddItem "ALIAN PLASTICS, S.A. DE C.V."
        .ComboBox13.AddItem "ALL STATE FASTENER CORPORATION"
        .ComboBox13.AddItem "ALL-RITE INDUSTRIES"
        .ComboBox13.AddItem "ALPHA INDUSTRY QUERETARO, S.A DE C.V."
        .ComboBox13.AddItem "ALPHYN INDUSTRIES INC."
        .ComboBox13.AddItem "ALPINE ELECTRONICS OF AMERICA"
        .ComboBox13.AddItem "ALTAK INC."
        .ComboBox13.AddItem "ALTO PRODUCTS CORP"
        .ComboBox13.AddItem "ALTRIDER LLC"
        .ComboBox13.AddItem "Amber Tooling CHINA LTD"
        .ComboBox13.AddItem "AMERICAN FURUKAWA, INC."
        .ComboBox13.AddItem "AMERICAN HOWA KENTUCKY INC."
        .ComboBox13.AddItem "AMERICAN MITSUBA CORPORATION"
        .ComboBox13.AddItem "AMERICAN MOLDED PRODUCTS"
        .ComboBox13.AddItem "AMERICAN PLASTIC MOLDING CORPORATION"
        .ComboBox13.AddItem "AMERICAN RECREATION PRODUCTS, INC."
        .ComboBox13.AddItem "AMERICAN STITCHCO INC"
        .ComboBox13.AddItem "AMI MANCHESTER"
        .ComboBox13.AddItem "AMITY MOLD COMPANY"
        .ComboBox13.AddItem "AMPHENOL ADRONICS"
        .ComboBox13.AddItem "ANA GLOBAL LLC"
        .ComboBox13.AddItem "ANIXTER FASTENERS"
        .ComboBox13.AddItem "ANIXTER"
        .ComboBox13.AddItem "ANSEI AMERICA, INC."
        .ComboBox13.AddItem "APEX SPRING & STAMPING CORP."
        .ComboBox13.AddItem "APG MEXICO AUTOMOTIVE PLASTICS GROUP MEXICO SA DE CV"
        .ComboBox13.AddItem "APTIV ELECTRIC SYSTEMS CO. LTD, JIANGMEN BRANCH"
        .ComboBox13.AddItem "APTIV SERVICES HONDURAS S. DE R.L. DE C.V."
        .ComboBox13.AddItem "APTIV SERVICES US, LLC"
        .ComboBox13.AddItem "ARCHEM AMERICA, INC."
        .ComboBox13.AddItem "ARCTERYX"
        .ComboBox13.AddItem "ARGUS CORPORATION"
        .ComboBox13.AddItem "ARKEL, INC."
        .ComboBox13.AddItem "ARNECOM INDUSTRIAS SA DE CV"
        .ComboBox13.AddItem "ARROTIN PLASTIC MATERIALS INDIANA"
        .ComboBox13.AddItem "ARTISAN MOLD & MACHINING CO."
        .ComboBox13.AddItem "ASPEN TECHNOLOGIES, INC."
        .ComboBox13.AddItem "ASPM"
        .ComboBox13.AddItem "ASSEMBLED PRODUCTS"
        .ComboBox13.AddItem "ASSOCIATED PACKAGING, INC."
        .ComboBox13.AddItem "ATC DRIVETRAIN LLC"
        .ComboBox13.AddItem "AURIA OLD FORT"
        .ComboBox13.AddItem "AURIA SOLUTIONS USA INC."
        .ComboBox13.AddItem "AUTECH JAPAN, INC."
        .ComboBox13.AddItem "AUTO ALLIANCE INTL INC"
        .ComboBox13.AddItem "AUTO PARTS MFG MISSISSIPPI"
        .ComboBox13.AddItem "AUTO VEHICLE PARTS"
        .ComboBox13.AddItem "AUTOLIV ASP, INC."
        .ComboBox13.AddItem "AUTOLIV NISSIN BRAKE SYSTEMS AMERICA LLC"
        .ComboBox13.AddItem "AUTOMATIC SPRING PRODUCTS CORP"
        .ComboBox13.AddItem "AUTOMOTIVE LIGHTING NORTH AMERICA"
        .ComboBox13.AddItem "AUTOMOTIVE VERITAS DE MEXICO S.A. DE C.V."
        .ComboBox13.AddItem "AUTONEUM BRASIL TEXTEIS ACUSTICOS LTDA"
        .ComboBox13.AddItem "AVANZAR INTERIOR TECHNOLOGIES DE MEXICO, S DE RL DE CV"
        .ComboBox13.AddItem "AVG OEAM INC."
        .ComboBox13.AddItem "AVG NORTH AMERICA, INC."
        .ComboBox13.AddItem "AXIOM PLASTICS INC."
        .ComboBox13.AddItem "AY MANUFACTURING"
        .ComboBox13.AddItem "BACKCOUNTRY SOLUTIONS, LLC"
        .ComboBox13.AddItem "BAIER & MICHELS USA, INC."
        .ComboBox13.AddItem "BAP"
        .ComboBox13.AddItem "BARTLETT"
        .ComboBox13.AddItem "BATES RUBBER, INC."
        .ComboBox13.AddItem "BATESVILLE TOOL"
        .ComboBox13.AddItem "BEACH MOLD"
        .ComboBox13.AddItem "BELLETECH CORP"
        .ComboBox13.AddItem "BELT-TECH PRODUCT INC."
        .ComboBox13.AddItem "BEND ALL AUTOMOTIVE INC."
        .ComboBox13.AddItem "BEST BOLT PRODUCTS INC"
        .ComboBox13.AddItem "BEYOND CLOTHING"
        .ComboBox13.AddItem "BH MOULD INDUSTRIAL LIMITED"
        .ComboBox13.AddItem "BLEVINS FABRICATION CORP"
        .ComboBox13.AddItem "BLUE STAR PLASTICS, INC."
        .ComboBox13.AddItem "BLUESMITHS LLC"
        .ComboBox13.AddItem "BM Industria Bergamasca Mobili S.P.A"
        .ComboBox13.AddItem "BORG WARNER"
        .ComboBox13.AddItem "BOS AUTOMOTIVE PRODUCTS CZ S.R.O."
        .ComboBox13.AddItem "BOS AUTOMOTIVE PRODUCTS IRAPUATO SA DE CV"
        .ComboBox13.AddItem "BOS GMBH & CO KG"
        .ComboBox13.AddItem "BOSSARD CANADA INC"
        .ComboBox13.AddItem "BOSSARD DE MEXICO SA DE CV"
        .ComboBox13.AddItem "BOSSARD DE MEXICO, S.A DE C.V"
        .ComboBox13.AddItem "BOSSARD DENMARK A/S"
        .ComboBox13.AddItem "BOSSARD DEUTSCHLAND GMBH"
        .ComboBox13.AddItem "BOSSARD FASTENING SOLUTIONS SHANGHAI CO., LTD."
        .ComboBox13.AddItem "BOSSARD NORTH AMERICA"
        .ComboBox13.AddItem "BOSSARD ONTARIO, INC."
        .ComboBox13.AddItem "BRADFORD INNOPACK"
        .ComboBox13.AddItem "BRIDGEWATER"
        .ComboBox13.AddItem "BRILLCAST INC."
        .ComboBox13.AddItem "C & D ZODIAC"
        .ComboBox13.AddItem "C & J TECH"
        .ComboBox13.AddItem "C & S PLASTICS"
        .ComboBox13.AddItem "CALSONICKANSEI MEXICANA, S.A. de C.V."
        .ComboBox13.AddItem "CAMPAMENTO S.A."
        .ComboBox13.AddItem "CAMRIG INC."
        .ComboBox13.AddItem "CAM-SLIDE MFG"
        .ComboBox13.AddItem "CAN-DO NATIONAL TAPE"
        .ComboBox13.AddItem "CAPCO LLC"
        .ComboBox13.AddItem "CARLEX GLASS COMPANY"
        .ComboBox13.AddItem "CARLEX GLASS OF INDIANA, INC."
        .ComboBox13.AddItem "CARPAK"
        .ComboBox13.AddItem "CASCADE DESIGNS, INC"
        .ComboBox13.AddItem "CASCADE ENGINEERING"
        .ComboBox13.AddItem "CATALINA COMPONENTS"
        .ComboBox13.AddItem "CCL DESIGN"
        .ComboBox13.AddItem "CENTRAL CAROLINA PRODUCTS, INC."
        .ComboBox13.AddItem "CENTURY FASTENERS CORP."
        .ComboBox13.AddItem "CEW ENTERPRISES, LLC"
        .ComboBox13.AddItem "CHALLENGE MFG. COMPANY LLC"
        .ComboBox13.AddItem "CHAMPION PLASTICS, INC."
        .ComboBox13.AddItem "CHANGAN FORD AUTOMOBILE COMPANY, LIMITED"
        .ComboBox13.AddItem "CHANGAN FORD MAZDA AUTOMOBILE"
        .ComboBox13.AddItem "CHANGSHA ADIENT AUTOMOTIVE COMPONENT CO, LTD."
        .ComboBox13.AddItem "CHASE PLASTIC SERVICES"
        .ComboBox13.AddItem "CHEWIE LABS, INC."
        .ComboBox13.AddItem "CHILA LLC., D/B/A MOCHIBRAND"
        .ComboBox13.AddItem "CHIYODA USA CORPORATION"
        .ComboBox13.AddItem "CHRISTOPHER P. SCHAFROTH"
        .ComboBox13.AddItem "CHUHATSU NORTH AMERICA, INC."
        .ComboBox13.AddItem "CIKAUXTO DE MEXICO"
        .ComboBox13.AddItem "CK TECHNOLOGIES, LLC"
        .ComboBox13.AddItem "CLARCORP INDUSTRIAL SALES"
        .ComboBox13.AddItem "Clariant"
        .ComboBox13.AddItem "CLARION TECHNOLOGIES INC."
        .ComboBox13.AddItem "CNI DULUTH LLC"
        .ComboBox13.AddItem "CNI, INC."
        .ComboBox13.AddItem "CO-EX-TEC INDUSTRIES"
        .ComboBox13.AddItem "COMMERCIAL VEHICLE GROUP, INC."
        .ComboBox13.AddItem "COMPONENT SUPPLY LLC"
        .ComboBox13.AddItem "COMPONENT TECHNOLOGIES INT'L INC."
        .ComboBox13.AddItem "COMPOSITE TECHNIQUES, INC."
        .ComboBox13.AddItem "COMTEL CORPORATION"
        .ComboBox13.AddItem "CONCORD TOOL & MANUFACTURING INC."
        .ComboBox13.AddItem "CONCOURS MOLD MEXICANA"
        .ComboBox13.AddItem "CONDUMEX, INC."
        .ComboBox13.AddItem "CONFORM AUTOMOTIVE"
        .ComboBox13.AddItem "CONFORMANCE FASTENERS"
        .ComboBox13.AddItem "CONSOLIDATED METCO, INC."
        .ComboBox13.AddItem "CONTINENTAL AUTOMOTIVE GUADALAJARA MEXICO S. DE R.L. C.V."
        .ComboBox13.AddItem "CONTINENTAL AUTOMOTIVE MEXICANA, S. DE R.L. C.V."
        .ComboBox13.AddItem "CONTINENTAL AUTOMOTIVE SYSTEMS, INC."
        .ComboBox13.AddItem "CONTINENTAL INDUSTRIES"
        .ComboBox13.AddItem "CONTINENTAL MANUFACTURING"
        .ComboBox13.AddItem "CONTINENTAL TEMIC ELECTRONICS PHILS.,INC."
        .ComboBox13.AddItem "CONTITECH FLUID DISTRIBUIDORA SA DE CV"
        .ComboBox13.AddItem "CONTITECH NORTH AMERICA, INC."
        .ComboBox13.AddItem "COOPER STANDARD AUTOMOTIVE BRASIL SEALING LTDA"
        .ComboBox13.AddItem "COOPER-STANDARD AUTOMOTIVE CESKA REPUBLIKA S.R.O."
        .ComboBox13.AddItem "CORPORACION MITSUBA DE MEXICO, S.A. DE C.V"
        .ComboBox13.AddItem "CORVAC COMPOSITES LLC - IN"
        .ComboBox13.AddItem "CORVAC COMPOSITES LLC- KY"
        .ComboBox13.AddItem "CORVAC COMPOSITES, LLC - MI"
        .ComboBox13.AddItem "COVAL MANUFACTURING S.A. DE C.V."
        .ComboBox13.AddItem "CRANE 1 SERVICES"
        .ComboBox13.AddItem "CRANE MERCHANDISING SYSTEMS"
        .ComboBox13.AddItem "CREATIVE LIQUID COATINGS INC."
        .ComboBox13.AddItem "CRISTALES INASTILLABLES DE MEXICO, S.A. DE C.V."
        .ComboBox13.AddItem "CRITERION TECHNOLOGY, INC."
        .ComboBox13.AddItem "CROWN PACKAGING CORP"
        .ComboBox13.AddItem "CS MANUFACTURING INC."
        .ComboBox13.AddItem "CST GMBH"
        .ComboBox13.AddItem "CUMBERLAND PLASTIC SYSTEM LLC"
        .ComboBox13.AddItem "CURTIS-MARUYASU AMERICA, INC."
        .ComboBox13.AddItem "CUSTOM MOLDED PRODUCTS LLC"
        .ComboBox13.AddItem "CUSTOMIZED MANUFACTURING AND ASSEMBLY"
        .ComboBox13.AddItem "CYTORI THERAPEUTICS"
        .ComboBox13.AddItem "D&N BENDING"
        .ComboBox13.AddItem "D.A. INC"
        .ComboBox13.AddItem "DAEHAN SOLUTION GEORGIA LLC"
        .ComboBox13.AddItem "DAEHAN SOLUTION NEVADA, LLC"
        .ComboBox13.AddItem "DAIEI AMERICA, INC."
        .ComboBox13.AddItem "DAIKYONISHIKAWA MEXICANA SA DE CV SP"
        .ComboBox13.AddItem "DAIKYONISHIKAWA USA INC."
        .ComboBox13.AddItem "DAIMAY AUTOMOTIVE INTERIOR S DE RL DE CV"
        .ComboBox13.AddItem "DAKKOTA INTEGRATED SYSTEMS"
        .ComboBox13.AddItem "DANNER DISTRIBUTION INC."
        .ComboBox13.AddItem "DAYSTAR CUT & SEW INC."
        .ComboBox13.AddItem "DDM"
        .ComboBox13.AddItem "DECATOR MOLD & TOOL"
        .ComboBox13.AddItem "DECATUR PLASTIC PRODUCTS, INC."
        .ComboBox13.AddItem "DECOPLAS S. A. DE C. V."
        .ComboBox13.AddItem "DECOPLAS SA DE CV"
        .ComboBox13.AddItem "DECOSTAR INDUSTRIES, INC."
        .ComboBox13.AddItem "DELPHI AUTOMOTIVE SYSTEMS, LLC"
        .ComboBox13.AddItem "DENSO AIR SYSTEMS DE MEXICO S. A. DE C. V."
        .ComboBox13.AddItem "DENSO AIR SYSTEMS DE MEXICO"
        .ComboBox13.AddItem "DENSO AIR SYSTEMS MICHIGAN INC"
        .ComboBox13.AddItem "DENSO MANUFACTURING CANADA, INC."
        .ComboBox13.AddItem "DENSO MANUFACTURING"
        .ComboBox13.AddItem "DENSO MANUFACTURING NORTH CAROLINA,INC."
        .ComboBox13.AddItem "DENSO MEXICO SA DE CV"
        .ComboBox13.AddItem "DENSO MEXICO, SA DE CV"
        .ComboBox13.AddItem "DENSO MFG TN INC - ATHENS"
        .ComboBox13.AddItem "DENSO MFG TN INC - MARYSVILLE"
        .ComboBox13.AddItem "DENSO MFG. ARKANSAS"
        .ComboBox13.AddItem "DEXTER STAMPING"
        .ComboBox13.AddItem "DICKTEN MASCH PLASTICS"
        .ComboBox13.AddItem "DISCO AUTOMOTIVE"
        .ComboBox13.AddItem "DIETECH TOOL & MFG., INC."
        .ComboBox13.AddItem "DISPENSING DYNAMICS INTERNATIONAL"
        .ComboBox13.AddItem "DIVERSATECH PLASTICS"
        .ComboBox13.AddItem "DIVERSITY - VUTEQ, LLC"
        .ComboBox13.AddItem "DIVERSITY - VUTEQ LLC MS"
        .ComboBox13.AddItem "DK MANUFACTURING FRAZEYSBURG"
        .ComboBox13.AddItem "DK MANUFACTURING LANCASTER, INC."
        .ComboBox13.AddItem "DLHBOWLES, INC."
        .ComboBox13.AddItem "DONALD HILTON"
        .ComboBox13.AddItem "DONG HEE CO., LTD"
        .ComboBox13.AddItem "DONG KWANG TECH"
        .ComboBox13.AddItem "DONGHEE KAUTEX LLC"
        .ComboBox13.AddItem "DORTEC INDUSTRIES"
        .ComboBox13.AddItem "DRAXLMAIER AUTOMOTIVE"
        .ComboBox13.AddItem "E & E MANUFACTURING"
        .ComboBox13.AddItem "EAGLE BEND"
        .ComboBox13.AddItem "EAKAS ARKANSAS CORPORATION"
        .ComboBox13.AddItem "EAST HK MOLDING COMPANY LTD"
        .ComboBox13.AddItem "EAST HAMILTON INDUSTRIES, INC."
        .ComboBox13.AddItem "ECHO ENGINEERING & PRODUCTION SUPPLIES INC."
        .ComboBox13.AddItem "ECHO NINER"
        .ComboBox13.AddItem "EDDIE BAUER"
        .ComboBox13.AddItem "EFC INTERNATIONAL"
        .ComboBox13.AddItem "EG INDUSTRIES CANADA"
        .ComboBox13.AddItem "EG INDUSTRIES CIRCLEVILLE"
        .ComboBox13.AddItem "EIMO TECHNOLOGIES, INC."
        .ComboBox13.AddItem "EIS FIBERCOATING, INC."
        .ComboBox13.AddItem "EISSMANN AUTOMOTIVE DETROIT DEVELOPMENT, LLC"
        .ComboBox13.AddItem "ELASTOMEROS TECNICOS MOLDEADOS INC."
        .ComboBox13.AddItem "ELCOM, INC."
        .ComboBox13.AddItem "ELEMATEC USA"
        .ComboBox13.AddItem "ELSA"
        .ComboBox13.AddItem "EMHART"
        .ComboBox13.AddItem "EMHART TEK"
        .ComboBox13.AddItem "EMMA HILL MANUFACTURING"
        .ComboBox13.AddItem "EMRICK PLASTICS"
        .ComboBox13.AddItem "ENDRIES INTERNATIONAL INC."
        .ComboBox13.AddItem "ENGINEERED APPAREL LTD."
        .ComboBox13.AddItem "ENGINEERED APPAREL, S.A. DE C.V."
        .ComboBox13.AddItem "ENGINEERED COMPONENT & SEAL"
        .ComboBox13.AddItem "ENGINEERED PARTS SOURCING INC."
        .ComboBox13.AddItem "ENGINEERED PLASTIC COMPONENTS"
        .ComboBox13.AddItem "ENNOVEA, LLC"
        .ComboBox13.AddItem "ENVISION AESC US LLC"
        .ComboBox13.AddItem "EQ SWIMWEAR"
        .ComboBox13.AddItem "EQUINOX, LTD."
        .ComboBox13.AddItem "ERIC SCOTT LEATHERS LTD"
        .ComboBox13.AddItem "ERLER INDUSTRIES INC."
        .ComboBox13.AddItem "ESON PRECISION INDUSTRY SINGAPORE PTE LTD"
        .ComboBox13.AddItem "EXCELL USA"
        .ComboBox13.AddItem "F & H"
        .ComboBox13.AddItem "F & P GEORGIA MFG., INC."
        .ComboBox13.AddItem "F & P MFG., INC."
        .ComboBox13.AddItem "F. PATRICK SZUSTAK"
        .ComboBox13.AddItem "FABRICA PORTUGUESA DE MOLDES PARA PLASTICOS LDA"
        .ComboBox13.AddItem "FALCON PLASTICS, INC."
        .ComboBox13.AddItem "FASTENER SUPPLY COMPANY"
        .ComboBox13.AddItem "FAURECIA AUTOMOTIVE POLSKA S.A."
        .ComboBox13.AddItem "FAURECIA AUTOMOTIVE SEATING"
        .ComboBox13.AddItem "FAURECIA INTERIORS SYSTEMS"
        .ComboBox13.AddItem "FAURECIA INTERIOR SYSTEMS, INC. TUSCALOOSA"
        .ComboBox13.AddItem "FAURECIA SISTEMAS AUTO DE MEXICO"
        .ComboBox13.AddItem "FAURECIA SISTEMAS AUTOMOTRICES DE MEXICO SA DE CV"
        .ComboBox13.AddItem "FEDERAL MOGUL SPG"
        .ComboBox13.AddItem "FENA APD"
        .ComboBox13.AddItem "FIC AMERICA"
        .ComboBox13.AddItem "FICOSA DO BRASIL LTDA."
        .ComboBox13.AddItem "FIH MEXICO INDUSTRY SA DE CV"
        .ComboBox13.AddItem "FIO AUTOMOTIVE CANADA"
        .ComboBox13.AddItem "FISCHER AUTOMOTIVE"
        .ComboBox13.AddItem "FLAMBEAU INC."
        .ComboBox13.AddItem "FLEETWOOD METAL INDUSTRIES"
        .ComboBox13.AddItem "FLEX N GATE COVINGTON"
        .ComboBox13.AddItem "FLEX N GATE TROY"
        .ComboBox13.AddItem "FLEXIBLE CIRCUIT TECHNOLOGIES, INC."
        .ComboBox13.AddItem "FLEX-N-GATE ALABAMA, LLC."
        .ComboBox13.AddItem "FLEX-N-GATE MEXICO PLASTICOS S. DE R. L. DE C.V."
        .ComboBox13.AddItem "FLEX-N-GATE OKLAHOMA LLC"
        .ComboBox13.AddItem "FLEX-N-GATE PLASTICS"
        .ComboBox13.AddItem "FLEXTRONICS INTERNATIONAL EUROPE BV"
        .ComboBox13.AddItem "FLORIDA PRODUCTION ENGINEERING"
        .ComboBox13.AddItem "FLORIDA PRODUCTION ENG-FL"
        .ComboBox13.AddItem "FOAM MOLDERS AND SPECIALTIES"
        .ComboBox13.AddItem "FORD ARGENTINA S.C.A"
        .ComboBox13.AddItem "FORD COMPONENT SALES, LLC."
        .ComboBox13.AddItem "FORD CUSTOMER SERVICE"
        .ComboBox13.AddItem "FORD ESPANA S.L."
        .ComboBox13.AddItem "FORD INDIA LIMITED"
        .ComboBox13.AddItem "FORD LIO HO MOTOR CO LTD"
        .ComboBox13.AddItem "FORD MOTOR COMPANY BRASIL LTDA"
        .ComboBox13.AddItem "FORD MOTOR COMPANY OF AUSTRALIA"
        .ComboBox13.AddItem "FORD MOTOR COMPANY SA DE CV"
        .ComboBox13.AddItem "FORD MOTOR COMPANY, SA DE CV"
        .ComboBox13.AddItem "FORD MOTOR COMPANY."
        .ComboBox13.AddItem "FORD RACING TECHNOLOGY"
        .ComboBox13.AddItem "FORD ROMANIA SA"
        .ComboBox13.AddItem "FORD-WERKE GMBH"
        .ComboBox13.AddItem "FORMA, LLC"
        .ComboBox13.AddItem "FRAENKISCHE INDUSTRIAL PIPES MEXICO SA DE CV"
        .ComboBox13.AddItem "FRAENKISCHE PIPE-SYSTEMS SHANGHAI CO., LTD."
        .ComboBox13.AddItem "FRANKLIN PRECISION INDUSTRY"
        .ComboBox13.AddItem "FREUDENBERG NOK"
        .ComboBox13.AddItem "FT PRECISION"
        .ComboBox13.AddItem "FUEL CELL SYSTEM MANUFACTURING"
        .ComboBox13.AddItem "FUEL TOTAL SYSTEMS KENTUCKY CORPORATION"
        .ComboBox13.AddItem "FUJI COMPONENT PARTS USA INC"
        .ComboBox13.AddItem "FUJIKURA AUTOMOTIVE AMERICA"
        .ComboBox13.AddItem "FURUKAWA AUTOMOTIVE SYSTEMS MEXICO SA DE CV"
        .ComboBox13.AddItem "FUTURIS AUTOMOTIVE CA LLC"
        .ComboBox13.AddItem "FUYAO AUTOMOTIVE NORTH AMERICA, INC."
        .ComboBox13.AddItem "FUYAO GLASS AMERICA INC."
        .ComboBox13.AddItem "G&B GLOBAL"
        .ComboBox13.AddItem "G.S.W. MANUFACTURING, INC."
        .ComboBox13.AddItem "GAJC"
        .ComboBox13.AddItem "GD COMPONENTS DE MEXICO SA DE CV"
        .ComboBox13.AddItem "GECOM"
        .ComboBox13.AddItem "GENERAL FASTENERS COMPANY"
        .ComboBox13.AddItem "GENERAL MOTORS"
        .ComboBox13.AddItem "GENERAL MOTORS DE ARGENTINA S.R.L."
        .ComboBox13.AddItem "GENERAL MOTORS DO BRASIL LTDA."
        .ComboBox13.AddItem "GENESIS CONCEPTS, LLC."
        .ComboBox13.AddItem "GENESIS PLASTICS & ENGINEERING"
        .ComboBox13.AddItem "GENESIS PLASTICS SOLUTIONS"
        .ComboBox13.AddItem "GENTEX CORPORATION"
        .ComboBox13.AddItem "GEXPRO"
        .ComboBox13.AddItem "GILL CORPORATION"
        .ComboBox13.AddItem "GIMME A PUTT LLC"
        .ComboBox13.AddItem "GL AUTOMOTIVE, LLC"
        .ComboBox13.AddItem "GLOBAL ENTERPRISES"
        .ComboBox13.AddItem "GLOBAL PLAS INC."
        .ComboBox13.AddItem "GLOBAL PLASTICS, INC."
        .ComboBox13.AddItem "GLOV ENTERPRISES, LLC"
        .ComboBox13.AddItem "GM DE MEXICO"
        .ComboBox13.AddItem "GOODWILL INDUSTRIES"
        .ComboBox13.AddItem "GOSSAMER GEAR INC."
        .ComboBox13.AddItem "GRAMMER AMERICAS"
        .ComboBox13.AddItem "GRAMMER AUTOMOTIVE PUEBLA"
        .ComboBox13.AddItem "GRAMMER INDUSTRIES"
        .ComboBox13.AddItem "GRAND RAPIDS CONTROLS"
        .ComboBox13.AddItem "GREAT LAKES ASSEMBLIES, LLC."
        .ComboBox13.AddItem "GREAT LAKES FASTENERS & SUPPLY CO."
        .ComboBox13.AddItem "GREEN INDUSTRIAL SUPPLY, INC."
        .ComboBox13.AddItem "GREENFIELD PRECISION PLASTICS, LLC."
        .ComboBox13.AddItem "GREENLEAF INDUSTRIES"
        .ComboBox13.AddItem "GREGORY MOUNTAIN PRODUCTS"
        .ComboBox13.AddItem "GROTE INDUSTRIES INC"
        .ComboBox13.AddItem "GRUPO ANTOLIN KENTUCKY, INC."
        .ComboBox13.AddItem "GRUPO ANTOLIN MISSOURI"
        .ComboBox13.AddItem "GRUPO ANTOLIN NORTH AMERICA"
        .ComboBox13.AddItem "GRUPO ANTOLIN PRIMERA"
        .ComboBox13.AddItem "GRUPO ANTOLIN SALTILLO S. DE R.L. DE C.V."
        .ComboBox13.AddItem "GRUPO ANTOLIN SILAO SA DE CV"
        .ComboBox13.AddItem "GRUPO ANTOLIN-SILAO, S.A. DE C.V."
        .ComboBox13.AddItem "GRUPO ANTOLIN ST CLAIR"
        .ComboBox13.AddItem "GRUPO MAQUILADOR DE XALAPA S.A. DE C.V."
        .ComboBox13.AddItem "GTR ENTERPRISES LLC"
        .ComboBox13.AddItem "GUARDIAN REPAIR & PARTS"
        .ComboBox13.AddItem "GUELPH MANUFACTURING GROUP INC."
        .ComboBox13.AddItem "GULF SHORE ASSEMBLIES, LLC."
        .ComboBox13.AddItem "H.S. DIE & ENGINEERING, INC."
        .ComboBox13.AddItem "HANCOCK MEDICAL INC"
        .ComboBox13.AddItem "HARADA INDUSTRY OF AMERICA, INC."
        .ComboBox13.AddItem "HARBOR ISLE PLASTICS LLC"
        .ComboBox13.AddItem "HARMAN/BECKER"
        .ComboBox13.AddItem "HARMINIE ENTERPRISE, INC."
        .ComboBox13.AddItem "HARMONY SYSTEMS AND SERVICE, INC."
        .ComboBox13.AddItem "HATCH STAMPING COMPANY"
        .ComboBox13.AddItem "HAWK PLASTICS LTD."
        .ComboBox13.AddItem "HAYAKAWA ELECTRONICS DE MEXICO S.A. DE C.V."
        .ComboBox13.AddItem "HAYASHI CANADA INC."
        .ComboBox13.AddItem "HAYASHI TELEMPU NORTH AMERICA CORP"
        .ComboBox13.AddItem "HAYASHI TELEMPU NORTH AMERICA CORP- CALF"
        .ComboBox13.AddItem "HBPO DE MEXICO"
        .ComboBox13.AddItem "HC QUERETARO, SA DE CV"
        .ComboBox13.AddItem "HEARTLAND AUTOMOTIVE"
        .ComboBox13.AddItem "HEBEI CHINAUST AUTOMOTIVE PLASTICS,LTD."
        .ComboBox13.AddItem "HELLA AUTOMOTIVE MEXICO S.A. DE C.V."
        .ComboBox13.AddItem "HELLA ELECTRONICS CORPORATION"
        .ComboBox13.AddItem "HERBERT E. ORR CO., INC."
        .ComboBox13.AddItem "HFI, LLC"
        .ComboBox13.AddItem "HI-CRAFT ENGINEERING"
        .ComboBox13.AddItem "HI-LEX DO BRASIL LTDA"
        .ComboBox13.AddItem "HIGHLANDS DIVERSIFIED SERVICES"
        .ComboBox13.AddItem "HI-LEX MEXICANA S. A. DE C. V."
        .ComboBox13.AddItem "HINO MOTORS CANADA"
        .ComboBox13.AddItem "HITACHI ASTEMO CAPAC, LLC"
        .ComboBox13.AddItem "HITACHI ASTEMO GREENFIELD, LLC"
        .ComboBox13.AddItem "HITACHI ASTEMO INDIANA, INC."
        .ComboBox13.AddItem "HITACHI ASTEMO MEXICO, S.A. DE C.V."
        .ComboBox13.AddItem "HITACHI ASTEMO OHIO MANUFACTURING, INC."
        .ComboBox13.AddItem "HITACHI AUTOMOTIVE SYSTEMS AMERICAS, INC."
        .ComboBox13.AddItem "HITACHI CABLE AMERICA, INC."
        .ComboBox13.AddItem "HI-TECH FASTENERS, INC."
        .ComboBox13.AddItem "HODELL-NATCO INDUSTRIES, INC."
        .ComboBox13.AddItem "HOLLINGSWORTH"
        .ComboBox13.AddItem "HONDA ACCESSORY AMERICA, LLC"
        .ComboBox13.AddItem "HONDA DE MEXICO PESO"
        .ComboBox13.AddItem "HONDA DE MEXICO SA DE CV CELAYA"
        .ComboBox13.AddItem "HONDA DE MEXICO, SA DE CV"
        .ComboBox13.AddItem "HONDA DEV & MFG OF AMERICA LLC"
        .ComboBox13.AddItem "HONDA MOTOR CO.,LTD."
        .ComboBox13.AddItem "HONDA OF AMERICA MFG INC AEP"
        .ComboBox13.AddItem "HONDA OF AMERICA MFG INC ELP"
        .ComboBox13.AddItem "HONDA OF AMERICA MFG INC IPS"
        .ComboBox13.AddItem "HONDA OF AMERICA MFG INC PMC"
        .ComboBox13.AddItem "HONDA OF INDIANA"
        .ComboBox13.AddItem "HONDA OF SOUTH CAROLINA INC"
        .ComboBox13.AddItem "HONDA POWER EQUIPMENT HPE"
        .ComboBox13.AddItem "HONDA PRECISION PARTS GEORGIA"
        .ComboBox13.AddItem "HONDA SUPPLY PARTS"
        .ComboBox13.AddItem "HONDA TRADING AMERICA"
        .ComboBox13.AddItem "HONDA TRADING DE MEXICO SA DE CV"
        .ComboBox13.AddItem "HONDA TRANSMISSION MFG."
        .ComboBox13.AddItem "HOOSIER MOLDED PRODUCTS INC."
        .ComboBox13.AddItem "HOPE GLOBAL"
        .ComboBox13.AddItem "HORIZON GLOBAL AMERICAS"
        .ComboBox13.AddItem "HORN"
        .ComboBox13.AddItem "HOSOUCHI MOLD CORPORATION"
        .ComboBox13.AddItem "HOUSE OF THREADS"
        .ComboBox13.AddItem "HOWA CANADA MANUFACTURING, INC."
        .ComboBox13.AddItem "HOWA MEXICO"
        .ComboBox13.AddItem "HOWA USA HOLDINGS INC."
        .ComboBox13.AddItem "HUDSON INDUSTRIES, INC"
        .ComboBox13.AddItem "HUNTER DOUGLAS CUSTOM SHUTTERS"
        .ComboBox13.AddItem "Hunter Industries Inc"
        .ComboBox13.AddItem "HUTCHINSON AUTOPARTES MEXICO, S.A. DE C.V."
        .ComboBox13.AddItem "HUTCHINSON TRANSFERENCIA DE FLUIDOS MEXICO"
        .ComboBox13.AddItem "HYDE EXPEDITION LLC"
        .ComboBox13.AddItem "HYPERLITE MOUNTAIN GEAR INCORPORATED"
        .ComboBox13.AddItem "IAC ALMA, LLC"
        .ComboBox13.AddItem "IAC ANNISTON LLC"
        .ComboBox13.AddItem "IAC CANADA ULC"
        .ComboBox13.AddItem "IAC DEARBORN"
        .ComboBox13.AddItem "IAC FREMONT LLC"
        .ComboBox13.AddItem "IAC GREENVILLE"
        .ComboBox13.AddItem "IAC GROUP GMBH"
        .ComboBox13.AddItem "IAC LEBANON LLC"
        .ComboBox13.AddItem "IAC MADISONVILLE KY"
        .ComboBox13.AddItem "IAC MENDON, LLC"
        .ComboBox13.AddItem "IAC SOUTHFIELD"
        .ComboBox13.AddItem "IAC SPRINGFIELD, LLC"
        .ComboBox13.AddItem "IAC WAUSEON, LLC"
        .ComboBox13.AddItem "IACNA GROUP, INC ARLINGTON"
        .ComboBox13.AddItem "IACNA MEXICO II S DE RL CV"
        .ComboBox13.AddItem "IACNA MEXICO OG"
        .ComboBox13.AddItem "IACNA SOFT TRIM CANADA"
        .ComboBox13.AddItem "IAC-STRASBURG"
        .ComboBox13.AddItem "ICON METAL FORMING LLC"
        .ComboBox13.AddItem "IEG PLASTICS, LLC"
        .ComboBox13.AddItem "II STANLEY COMPANY INC"
        .ComboBox13.AddItem "ILPEA S RL DE CV"
        .ComboBox13.AddItem "IMA"
        .ComboBox13.AddItem "IMA DETROIT"
        .ComboBox13.AddItem "INABATA AMERICA CORP."
        .ComboBox13.AddItem "INABATA MEXICO SA DE CV"
        .ComboBox13.AddItem "INDEPENDENT II,LLC"
        .ComboBox13.AddItem "INDIANA MARUJUN, LLC"
        .ComboBox13.AddItem "INDUSTRIA DE ASIENTO SUPERIOR"
        .ComboBox13.AddItem "INDUSTRIA DE ASIENTO SUPERIOR SA DE CV"
        .ComboBox13.AddItem "INDUSTRIAL CONVERTING CO"
        .ComboBox13.AddItem "INDUSTRIAL TECH SERVICES, INC."
        .ComboBox13.AddItem "INDUSTRIAS BM DE MEXICO SA DE CV"
        .ComboBox13.AddItem "INDUSTRIAS CAZEL"
        .ComboBox13.AddItem "INDUSTRIAS DMU, S. A. DE C. V."
        .ComboBox13.AddItem "INDUSTRIAS MANGOTEX LTDA"
        .ComboBox13.AddItem "INDUSTRIAS TRICON DE MEXICO"
        .ComboBox13.AddItem "INDUSTRIE ILPEA ESPANA, S.A"
        .ComboBox13.AddItem "INDUSTRY PRODUCTS CO"
        .ComboBox13.AddItem "INDUSTRY PRODUCTS CO."
        .ComboBox13.AddItem "INFOCASE INC"
        .ComboBox13.AddItem "INJEX"
        .ComboBox13.AddItem "INNERTECH"
        .ComboBox13.AddItem "INNOTEC"
        .ComboBox13.AddItem "INOAC EXTERIOR PRODUCTS, LLC."
        .ComboBox13.AddItem "INOAC EXTERIOR SYSTEMS INC."
        .ComboBox13.AddItem "INOAC GROUP NORTH AMERICA"
        .ComboBox13.AddItem "INOAC INTERIOR SYSTEMS LP"
        .ComboBox13.AddItem "INOAC SISTEMAS EXTERIORES, SA DE CV"
        .ComboBox13.AddItem "INPAQ TECHNOLOGY SUZHOU CO., LTD."
        .ComboBox13.AddItem "INSTASET PLASTICS CO. LLC"
        .ComboBox13.AddItem "INT AMERICA, LLC"
        .ComboBox13.AddItem "INTERLINK AUTOMOTIVE LLC"
        .ComboBox13.AddItem "INTERNATIONAL MOLD CORPORATION"
        .ComboBox13.AddItem "INTERTEC SYSTEMS"
        .ComboBox13.AddItem "INTERTEX TRADING CORP"
        .ComboBox13.AddItem "INTEVA PRODUCTS LLC"
        .ComboBox13.AddItem "INTEVA PRODUCTS, LLC"
        .ComboBox13.AddItem "INTEVA VANDALIA ENGINEERING CENTER"
        .ComboBox13.AddItem "INZI CONTROLS ALABAMA INC."
        .ComboBox13.AddItem "IRAPUATO PROPERTY AND ASSETS S DE RL DE CV"
        .ComboBox13.AddItem "IRVIN AUTOMOTIVE PRODUCTS"
        .ComboBox13.AddItem "ISGO NA LLC"
        .ComboBox13.AddItem "ISHMAEL PRECISION TOOL CORPORATION"
        .ComboBox13.AddItem "ISRINGHAUSEN GMBH & CO. KG"
        .ComboBox13.AddItem "ISUZU"
        .ComboBox13.AddItem "ITC INC."
        .ComboBox13.AddItem "ITW DELTAR BODY AND INTERIOR"
        .ComboBox13.AddItem "ITW DELTAR FUEL SYSTEMS"
        .ComboBox13.AddItem "ITW DELTAR IPAC"
        .ComboBox13.AddItem "ITW MOTION - US"
        .ComboBox13.AddItem "JAC PRODUCTS"
        .ComboBox13.AddItem "JAC PRODUCTS PORTUGAL"
        .ComboBox13.AddItem "JACKSON PLASTICS OPERATIONS"
        .ComboBox13.AddItem "JACKSON PLASTICS INC"
        .ComboBox13.AddItem "JAE OREGON"
        .ComboBox13.AddItem "JAG MANUFACTURING INC."
        .ComboBox13.AddItem "JASPER RUBBER PRODUCTS, INC."
        .ComboBox13.AddItem "JATCO MEXICO, S.A DE C.V"
        .ComboBox13.AddItem "JATCO MEXICO, S.A. DE C.V."
        .ComboBox13.AddItem "JAY INDUSTRIES INC"
        .ComboBox13.AddItem "JAY PLASTICS A DIV OF JAY INDUSTRIES, INC."
        .ComboBox13.AddItem "JBC TECHNOLOGIES,INC."
        .ComboBox13.AddItem "JCIM, LLC"
        .ComboBox13.AddItem "JEFFERSON ELORA CORPORATION"
        .ComboBox13.AddItem "JET ELECTRIC"
        .ComboBox13.AddItem "JOCK-JEWELRY"
        .ComboBox13.AddItem "JONES PLASTIC & ENGINEERING"
        .ComboBox13.AddItem "JOYSON SAFETY SYSTEMS ACQUISITION LLC"
        .ComboBox13.AddItem "JR MANUFACTURING INC"
        .ComboBox13.AddItem "JTEKT COLUMN SYSTEMS NORTH AMERICA CORPORATION"
        .ComboBox13.AddItem "JULIE DRISCOLL"
        .ComboBox13.AddItem "JVIS USA, LLC"
        .ComboBox13.AddItem "K&S WIRING SYSTEMS"
        .ComboBox13.AddItem "KAMCO"
        .ComboBox13.AddItem "KANE-M INC"
        .ComboBox13.AddItem "KANTUS"
        .ComboBox13.AddItem "KANUK INC"
        .ComboBox13.AddItem "KASAI NORTH AMERICA, INC."
        .ComboBox13.AddItem "KASAI NORTH AMERICA,INC. ALABAMA DIVISION"
        .ComboBox13.AddItem "KATABATIC GEAR LLC"
        .ComboBox13.AddItem "KATAYAMA AMERICAN COMPANY"
        .ComboBox13.AddItem "KAUTEX CHONGQING PLASTIC TECHNOLOGY CO., LTD."
        .ComboBox13.AddItem "KAUTEX GUANGZHOU PLASTIC TECHNOLOGY CO., LTD."
        .ComboBox13.AddItem "KAUTEX SHANGHAI"
        .ComboBox13.AddItem "KAUTEX JAPAN CORPORATION"
        .ComboBox13.AddItem "KAUTEX TEXTRON DO BRASIL LTDA"
        .ComboBox13.AddItem "KAUTEX TEXTRON DE MEXICO"
        .ComboBox13.AddItem "KAUTEX TEXTRON GMBH & CO."
        .ComboBox13.AddItem "KAUTEX TEXTRON GMBH & CO. KG"
        .ComboBox13.AddItem "KAUTEX TEXTRON IBERICA S.L."
        .ComboBox13.AddItem "KAWASAKI TENNESSEE, INC."
        .ComboBox13.AddItem "KB COMPONENTS CANADA INC"
        .ComboBox13.AddItem "KENDRICK PLASTICS"
        .ComboBox13.AddItem "KEY MANUFACTURING, LLC"
        .ComboBox13.AddItem "KEY TRONIC CORP"
        .ComboBox13.AddItem "KI USA CORP"
        .ComboBox13.AddItem "KINESIS PHOTO GEAR"
        .ComboBox13.AddItem "KINUGAWA FABRICACAO, IMPORTACAO E EXPORTACAO DE PECAS AUTOMOTIVAS LTDA."
        .ComboBox13.AddItem "KITTYHAWK MOLDING"
        .ComboBox13.AddItem "KNOT TECHNOLOGY, INC."
        .ComboBox13.AddItem "KNT ASSOCIATES"
        .ComboBox13.AddItem "KOKATAT WATERSPORTS WEAR"
        .ComboBox13.AddItem "KOLLER GROUP MEXICO"
        .ComboBox13.AddItem "KOLLER-CRAFT SOUTH"
        .ComboBox13.AddItem "KOSTAL MEXICANA SA DE CV"
        .ComboBox13.AddItem "KOTOBUKIYA TREVES DE MEXICO SA DE CV"
        .ComboBox13.AddItem "KROMBERG & SCHUBERT MEXICO S. EN C."
        .ComboBox13.AddItem "KSR INTERNATIONAL CO."
        .ComboBox13.AddItem "KTNA INC."
        .ComboBox13.AddItem "KUMI"
        .ComboBox13.AddItem "KUMI ALABAMA"
        .ComboBox13.AddItem "KURI TEC"
        .ComboBox13.AddItem "KYOSAN DENSO MFG. KY LLC"
        .ComboBox13.AddItem "KYOWA AMERICA CORPORATION"
        .ComboBox13.AddItem "L EQUIPE MONTEUR SA"
        .ComboBox13.AddItem "L&W"
        .ComboBox13.AddItem "L.L. BEAN INC."
        .ComboBox13.AddItem "LACKS TRIM SYSTEMS"
        .ComboBox13.AddItem "LAKELAND FINISHING"
        .ComboBox13.AddItem "LAKEPARK INDUSTRIES, INC."
        .ComboBox13.AddItem "LAKESIDE PLASTICS LIMITED"
        .ComboBox13.AddItem "LEAR - EL PASO"
        .ComboBox13.AddItem "LEAR ARLINGTON"
        .ComboBox13.AddItem "LEAR AUTO SYSTEMS CHONGQING CO., LTD. METALS"
        .ComboBox13.AddItem "LEAR AUTOMOTIVE THAILAND CO., LTD."
        .ComboBox13.AddItem "LEAR AUTOMOTIVE METALS WUHAN CO., LTD."
        .ComboBox13.AddItem "LEAR CHANGAN CHONGQING AUTOMOTIVE SYSTEMS CO.,LTD"
        .ComboBox13.AddItem "LEAR CORP - AJAX"
        .ComboBox13.AddItem "LEAR CORP - OSHAWA"
        .ComboBox13.AddItem "LEAR CORP - WHITBY"
        .ComboBox13.AddItem "LEAR CORP SSD - MASON"
        .ComboBox13.AddItem "LEAR CORP SSD - NORTH AMERICAN DIV"
        .ComboBox13.AddItem "LEAR CORP SSD - RAMOS"
        .ComboBox13.AddItem "LEAR CORP SSD - ROCHESTER HILLS"
        .ComboBox13.AddItem "LEAR CORP SSDI - SALTILLO"
        .ComboBox13.AddItem "LEAR CORP SSDI- ROSCOMMON"
        .ComboBox13.AddItem "LEAR CORPORATION - ALABAMA"
        .ComboBox13.AddItem "LEAR CORPORATION FLINT"
        .ComboBox13.AddItem "LEAR CORPORATION GMBH"
        .ComboBox13.AddItem "LEAR CORPORATION GRATIOT"
        .ComboBox13.AddItem "Lear Corporation Headquarters"
        .ComboBox13.AddItem "LEAR CORPORATION ITALIA S.R.L A SOCIO UNICO"
        .ComboBox13.AddItem "LEAR CORPORATION LOUISVILLE"
        .ComboBox13.AddItem "LEAR CORPORATION POLAND"
        .ComboBox13.AddItem "LEAR CORPORATION SSD DUNCAN"
        .ComboBox13.AddItem "LEAR CORPORATION-TUSCALOOSA"
        .ComboBox13.AddItem "LEAR GM SEATING-CONNER STREET"
        .ComboBox13.AddItem "LEAR GRAND PRAIRIE"
        .ComboBox13.AddItem "LEAR MEXICAN SEATING CORP - TRIM DIV"
        .ComboBox13.AddItem "LEAR MEXICAN SEATING CORP ARTEAGA"
        .ComboBox13.AddItem "LEAR MEXICAN SEATING CORP-HERMOSILLO"
        .ComboBox13.AddItem "LEAR MEXICAN SEATING CORP-P.NEGRAS"
        .ComboBox13.AddItem "LEAR MEXICAN SEATING CORP-SAN LUIS"
        .ComboBox13.AddItem "LEAR MEXICAN SEATING CORP-SILAO"
        .ComboBox13.AddItem "LEAR MEXICAN SEATING CORP-TOLUCA"
        .ComboBox13.AddItem "LEAR MEXICAN SEATING-MONCLOVA"
        .ComboBox13.AddItem "LEAR TRIM OPERATIONS - PIEDRAS"
        .ComboBox13.AddItem "LEAR-MORRISTOWN-TN"
        .ComboBox13.AddItem "LEATHERMAN TOOL GROUP INC"
        .ComboBox13.AddItem "LEHMANN-PETERSON"
        .ComboBox13.AddItem "LENS MOLD"
        .ComboBox13.AddItem "LEON PLASTICS INC."
        .ComboBox13.AddItem "LEXAMAR CORP"
        .ComboBox13.AddItem "LEXPLASTICS, LLC"
        .ComboBox13.AddItem "LISTOWELL TECHNOLOGY"
        .ComboBox13.AddItem "LM MANUFACTURING"
        .ComboBox13.AddItem "LN OF AMERICA INC."
        .ComboBox13.AddItem "LOGISTIK UNICORP INC."
        .ComboBox13.AddItem "LONGARM PRODUCTS INC. DBA: MY AD CANADA"
        .ComboBox13.AddItem "LORDSTOWN SEATING SYSTEMS"
        .ComboBox13.AddItem "LOST ARROW CORP"
        .ComboBox13.AddItem "LUCENT POLYMERS"
        .ComboBox13.AddItem "LUNA SANDAL CO. LLC"
        .ComboBox13.AddItem "LUND INTERNATIONAL"
        .ComboBox13.AddItem "LUNKETEC DE MEXICO"
        .ComboBox13.AddItem "LUNKETEC DE MEXICO S.A. DE C.V."
        .ComboBox13.AddItem "MACA PLASTICS, INC."
        .ComboBox13.AddItem "MADISON FASTENERS, LLC"
        .ComboBox13.AddItem "MAGNA ASSEMBLY SYSTEMS DE MEXICO"
        .ComboBox13.AddItem "MAGNA CLOSURES DE MEXICO"
        .ComboBox13.AddItem "MAGNA CLOSURES DO BRASIL"
        .ComboBox13.AddItem "MAGNA ELECTRONICS, INC."
        .ComboBox13.AddItem "MAGNA EXTERIORS BELVIDERE"
        .ComboBox13.AddItem "MAGNA EXTERIORS TOLUCA 'MAST"""
        .ComboBox13.AddItem "MAGNA MIRROR SYSTEMS MONTERREY"
        .ComboBox13.AddItem "MAGNA MIRRORS NORTH AMERICA"
        .ComboBox13.AddItem "MAGNA MIRRORS OF AMERICA, INC."
        .ComboBox13.AddItem "MAGNA SEALING AND GLASS"
        .ComboBox13.AddItem "MAGNA SEATING AUBURN HILLS"
        .ComboBox13.AddItem "MAGNA SEATING COLUMBUS"
        .ComboBox13.AddItem "MAGNA SEATING DETROIT"
        .ComboBox13.AddItem "MAGNA SEATING OF AMERICA"
        .ComboBox13.AddItem "MAGNA SEATING-CHATTANOOGA"
        .ComboBox13.AddItem "MAGNUSON PRODUCTS LLC"
        .ComboBox13.AddItem "MAHLE BEHR MT. STERLING, INC."
        .ComboBox13.AddItem "MAHLE FILTER SYSTEMS"
        .ComboBox13.AddItem "MAHLE SISTEMAS DE FILTRACION"
        .ComboBox13.AddItem "MAHLE SISTEMAS DE FILTRACION DE MEXICO S.A. DE C.V."
        .ComboBox13.AddItem "Majestics Plastics"
        .ComboBox13.AddItem "MANTALINE CORPORATION"
        .ComboBox13.AddItem "MARCO POLO INTERNATIONAL, LLC."
        .ComboBox13.AddItem "MARELLI MEXICANA, SA DE CV"
        .ComboBox13.AddItem "MARELLI NORTH AMERICA,INC."
        .ComboBox13.AddItem "MARIAH OF OHIO LLC"
        .ComboBox13.AddItem "MARNE PLASTICS"
        .ComboBox13.AddItem "MARTINREA AUTOMOTIVE STRUCTURES S. DE R.L DE C.V."
        .ComboBox13.AddItem "MARWOOD METAL FABRICATION"
        .ComboBox13.AddItem "MASTERGUARD"
        .ComboBox13.AddItem "MASTERMOLDING"
        .ComboBox13.AddItem "MATCOR AUTOMOTIVE"
        .ComboBox13.AddItem "MATSU MANUFACTURING BARRIE INC."
        .ComboBox13.AddItem "MAXEY INDUSTRIES"
        .ComboBox13.AddItem "MAYCO AUTOMOTIVE INTERNATIONAL S. DE R.L. DE C.V."
        .ComboBox13.AddItem "MAYCO INTERNATIONAL LLC"
        .ComboBox13.AddItem "MAYFAIR PLASTICS, INC."
        .ComboBox13.AddItem "MAZDA MOTOR CORPORATION"
        .ComboBox13.AddItem "MCKECHNIE VEHICLE COMPONENTS"
        .ComboBox13.AddItem "MCMINNVILLE TOOL & DIE, INC."
        .ComboBox13.AddItem "MCMURRAY FABRICS"
        .ComboBox13.AddItem "MEC"
        .ComboBox13.AddItem "MEIKI CORPORATION"
        .ComboBox13.AddItem "MEIWA INDUSTRY NORTH AMERICA, INC."
        .ComboBox13.AddItem "MELANZANA"
        .ComboBox13.AddItem "MERGON CORP"
        .ComboBox13.AddItem "METAL AND PLASTIC PAINT SOLUTIONS SA DE CV"
        .ComboBox13.AddItem "METALSA S.A DE C.V."
        .ComboBox13.AddItem "METALSA S DE RL"
        .ComboBox13.AddItem "MG INTERNATIONAL"
        .ComboBox13.AddItem "MID-STATES BOLT & SCREW CO."
        .ComboBox13.AddItem "MIDWAY PRODUCTS GROUP, INC."
        .ComboBox13.AddItem "MIDWEST MOLDING"
        .ComboBox13.AddItem "MIGAVID INDUSTRALES SA DE CV"
        .ComboBox13.AddItem "MILCO INDUSTRIES, INC."
        .ComboBox13.AddItem "MINEBEA ACCESSSOLUTIONS USA INC."
        .ComboBox13.AddItem "MINTH MEXICO COATINGS S.A. DE C.V."
        .ComboBox13.AddItem "MINTH MEXICO, S.A. DE C.V."
        .ComboBox13.AddItem "MISC PRODUCTS, INC."
        .ComboBox13.AddItem "MITCHELL PLASTICS"
        .ComboBox13.AddItem "MITCHELL PLASTICS, A DIV OF ULTRA MFG USA, INC."
        .ComboBox13.AddItem "MITCHELL PLASTICS, A DIV OF ULTRA MFG LTD"
        .ComboBox13.AddItem "MITSUBISHI MOTOR COMPANY"
        .ComboBox13.AddItem "MITSUBISHI MOTORS AUSTRALIA"
        .ComboBox13.AddItem "MITSUI KINZOKU ACT MEXICANA, S.A. DE C.V."
        .ComboBox13.AddItem "ML INDUSTRIES"
        .ComboBox13.AddItem "MM PLASTICS, LLC"
        .ComboBox13.AddItem "MME GROUP"
        .ComboBox13.AddItem "MOHR ENGINEERING, INC."
        .ComboBox13.AddItem "MOLLERTECH LLC"
        .ComboBox13.AddItem "MOLTEN AUTOMOTIVE DE MEXICO SA DE CV"
        .ComboBox13.AddItem "MOLTEN CORP"
        .ComboBox13.AddItem "MONTAPLAST OF NORTH AMERICA"
        .ComboBox13.AddItem "MORIDEN AMERICA INC"
        .ComboBox13.AddItem "MORIROKU AMERICA, INC."
        .ComboBox13.AddItem "MORIROKU TECH NA-GREENVILLE PLANT"
        .ComboBox13.AddItem "MORIROKU TECH NA-RAINSVILLE PLANT"
        .ComboBox13.AddItem "MORIROKU TECHNOLOGY DE MEXICO SA DE CV"
        .ComboBox13.AddItem "MORITO SCOVILL MEXICO"
        .ComboBox13.AddItem "MOTHERSON SUMI SYSTEM LIMITED"
        .ComboBox13.AddItem "MOTORES Y APARATOS ELECTRICOS"
        .ComboBox13.AddItem "MOTUS INTEGRATED TECHNOLOGIES"
        .ComboBox13.AddItem "MOUNTAIN WEAR CONFECCOES"
        .ComboBox13.AddItem "MPC INC"
        .ComboBox13.AddItem "MPI MACHINERY AND DESIGN LLC"
        .ComboBox13.AddItem "MSB"
        .ComboBox13.AddItem "MT BORAH"
        .ComboBox13.AddItem "MULTI-FORM PLASTICS INC"
        .ComboBox13.AddItem "MURAKAMI MANUFACTURING U.S.A., INC."
        .ComboBox13.AddItem "MUSKEGON CASTING CORP"
        .ComboBox13.AddItem "MVA STRATFORD INC."
        .ComboBox13.AddItem "MW MONROE PLASTICS, INC."
        .ComboBox13.AddItem "MWW AUTOMOTIVE CORP/COLORTEK"
        .ComboBox13.AddItem "MYSTERY RANCH, LTD"
        .ComboBox13.AddItem "NACOM CORPORATION"
        .ComboBox13.AddItem "NAGASE AMERICA LLC"
        .ComboBox13.AddItem "NASCOT"
        .ComboBox13.AddItem "NATIONAL CYCLE, INC."
        .ComboBox13.AddItem "NATIONAL ENGINEERED FASTENERS"
        .ComboBox13.AddItem "NEATON AUTO MEXICANA SA DE CV"
        .ComboBox13.AddItem "NEOCON"
        .ComboBox13.AddItem "NEWMAN TECHNOLOGY INC"
        .ComboBox13.AddItem "NEWMAN TECHNOLOGY OF ALABAMA,INC."
        .ComboBox13.AddItem "NEXGEN MOLD AND TOOL, INC."
        .ComboBox13.AddItem "NHK SEATING OF AMERICA, INC."
        .ComboBox13.AddItem "NI AUTOWINDOW"
        .ComboBox13.AddItem "NICHIRIN TENNESSEE INC."
        .ComboBox13.AddItem "NICHIRIN-FLEX USA INC."
        .ComboBox13.AddItem "NIFAST"
        .ComboBox13.AddItem "NIFAST CORPORATION GEORGIA"
        .ComboBox13.AddItem "NIFCO HK"
        .ComboBox13.AddItem "NIFCO HUBEI CO., LTD"
        .ComboBox13.AddItem "NIFCO THAILAND CO., LTD."
        .ComboBox13.AddItem "NIFCO THAILAND CO., LTD. USD"
        .ComboBox13.AddItem "NIFCO CENTRAL MEXICO S DE RL DE CV"
        .ComboBox13.AddItem "NIFCO DE MEXICO"
        .ComboBox13.AddItem "NIFCO GERMANY GMBH"
        .ComboBox13.AddItem "NIFCO INC"
        .ComboBox13.AddItem "NIFCO INC. TECHNOLOGY CENTER II"
        .ComboBox13.AddItem "NIFCO INDIA PVT. LTD."
        .ComboBox13.AddItem "NIFCO INDONESIA"
        .ComboBox13.AddItem "NIFCO KOREA INC."
        .ComboBox13.AddItem "NIFCO KOREA USA INC"
        .ComboBox13.AddItem "NIFCO KTS GMBH"
        .ComboBox13.AddItem "NIFCO KTW AMERICA CORPORATION"
        .ComboBox13.AddItem "NIFCO MANUFACTURING MALAYSIA SDN.BHD."
        .ComboBox13.AddItem "NIFCO POLAND"
        .ComboBox13.AddItem "NIFCO PRODUCTS ESPANA"
        .ComboBox13.AddItem "NIFCO SOUTH INDIA MANUFACTURING PVT. LTD."
        .ComboBox13.AddItem "NIFCO STAFFING SERVICE S DE RL DE CV"
        .ComboBox13.AddItem "NIFCO TECHNOLOGY DEVELOPMENT CENTER"
        .ComboBox13.AddItem "NIFCO UK LIMITED"
        .ComboBox13.AddItem "NIFCO VIETNAM LIMITED"
        .ComboBox13.AddItem "NIHON PLAST MEXICANA, S.A. DE C.V."
        .ComboBox13.AddItem "NIHON PLAST MEXICANA, S.A. DE C.V. NCM"
        .ComboBox13.AddItem "NILE AUTOMOTIVE GROUP TENNESSEE LLC."
        .ComboBox13.AddItem "NISHIKAWA SEALING SYSTEMS MEXICO SA DE CV"
        .ComboBox13.AddItem "NISSAN CORPORATION"
        .ComboBox13.AddItem "NISSAN EXPORTS DE MEXICO, S. DE R.L. DE C.V."
        .ComboBox13.AddItem "NISSAN MEXICANA PCC"
        .ComboBox13.AddItem "Nissan Mexicana SA de CV 3M NO USAR"
        .ComboBox13.AddItem "NISSAN MEXICANA SERVICE"
        .ComboBox13.AddItem "NISSAN MEXICANA TECHNICAL CENTER"
        .ComboBox13.AddItem "NISSAN MEXICANA X11M ARIES NO USAR"
        .ComboBox13.AddItem "NISSAN MEXICANA, S.A. DE C.V PESO"
        .ComboBox13.AddItem "NISSAN PDC"
        .ComboBox13.AddItem "NISSAN NORTH AMERICA, INC."
        .ComboBox13.AddItem "NISSAN TECHNICAL CENTER NORTH"
        .ComboBox13.AddItem "NISSAN TRADING CORPORATION AMERICAS"
        .ComboBox13.AddItem "NISSEN CHEMITEC AMERICA"
        .ComboBox13.AddItem "Nissen Chemitec Mexico, SA de CV"
        .ComboBox13.AddItem "NISSIN BRAKE GEORGIA"
        .ComboBox13.AddItem "NISSINKAKOU MEXICANA, S.A. DE C.V."
        .ComboBox13.AddItem "NITTA MOORE MEXICO, S. DE R.L DE C.V"
        .ComboBox13.AddItem "NOBEL AUTOMOTIVE MEXICO"
        .ComboBox13.AddItem "NORPLAS INDUSTRIES INC."
        .ComboBox13.AddItem "NORTH AMERICAN LIGHTING"
        .ComboBox13.AddItem "NORTH AMERICAN LIGHTING MEXICO, S.A DE C.V."
        .ComboBox13.AddItem "NOVARES"
        .ComboBox13.AddItem "NOVARES US ENGINE COMPONENTS,INC."
        .ComboBox13.AddItem "NOVEM CAR INTERIOR DESIGN"
        .ComboBox13.AddItem "NYLONCRAFT"
        .ComboBox13.AddItem "NYX INC."
        .ComboBox13.AddItem "NYX MEXICO PLASTICS S DE RL DE CV"
        .ComboBox13.AddItem "O & S CALIFORNIA, INC."
        .ComboBox13.AddItem "OHASHI"
        .ComboBox13.AddItem "OHIO TRANSMISSION CORPORATION"
        .ComboBox13.AddItem "OJI INTERTECH"
        .ComboBox13.AddItem "OKAYA U.S.A, INC."
        .ComboBox13.AddItem "OMNI"
        .ComboBox13.AddItem "OPTIMAS OE SOLUTIONS HOLDING, LLC"
        .ComboBox13.AddItem "ORC INDUSTRIES INC."
        .ComboBox13.AddItem "ORION INNOVATIONS, LLC"
        .ComboBox13.AddItem "OROTEX CORPORATION"
        .ComboBox13.AddItem "OTSCON INC."
        .ComboBox13.AddItem "OTSCON MEXICO MANUFACTURING, S.A. DE C.V."
        .ComboBox13.AddItem "OTTE GEAR LLC"
        .ComboBox13.AddItem "PACIFIC COAST"
        .ComboBox13.AddItem "PACIFIC INSIGHT ELECTRONICS CORP."
        .ComboBox13.AddItem "PACIFIC MANUFACTURING OHIO INC."
        .ComboBox13.AddItem "PACIFIC PLASTIC TECHNOLOGY"
        .ComboBox13.AddItem "PADDLERS SUPPLY"
        .ComboBox13.AddItem "PAK-RITE INDUSTRIES, INC."
        .ComboBox13.AddItem "PALM PLASTICS"
        .ComboBox13.AddItem "PANASONIC AUTOMOTIVE SYSTEMS COMPANY OF AMERICA"
        .ComboBox13.AddItem "PANASONIC AUTOMOTIVE SYSTEMS DE MEXICO S.A. DE C.V."
        .ComboBox13.AddItem "PANASONIC SHIKOKU ELECTRONICS"
        .ComboBox13.AddItem "PANGAEA, LTD."
        .ComboBox13.AddItem "PAPP PLASTICS & DISTRIBUTING LTD"
        .ComboBox13.AddItem "PAR 4 PLASTICS, INC."
        .ComboBox13.AddItem "Parker Corporation Mexicana S.A. de C.V."
        .ComboBox13.AddItem "PASA PANASONIC"
        .ComboBox13.AddItem "PATTON INDUSTRIAL PRODUCTS"
        .ComboBox13.AddItem "PEGASUS PACKAGING"
        .ComboBox13.AddItem "PENDA CORPORATION"
        .ComboBox13.AddItem "PENINSULA COMPONENTS INC"
        .ComboBox13.AddItem "PENSTONE INC"
        .ComboBox13.AddItem "Perfection Components"
        .ComboBox13.AddItem "PERFORMANCE ASSEMLBY SOLUTIONS"
        .ComboBox13.AddItem "PERLANE SALES INC."
        .ComboBox13.AddItem "PHILLIPS AND TEMRO"
        .ComboBox13.AddItem "PHILLIPS DIVERSIFIED MANUFACTURING, INC."
        .ComboBox13.AddItem "PHOENIX ASSEMBLY LLC"
        .ComboBox13.AddItem "PILKINGTON AUTOMOTIVE ARGENTINA S.A."
        .ComboBox13.AddItem "PILKINGTON BRASIL"
        .ComboBox13.AddItem "PINTURA Y ENSAMBLES DE MEXICO"
        .ComboBox13.AddItem "PIOLAX CORPORATION"
        .ComboBox13.AddItem "PISTON AUTOMOTIVE"
        .ComboBox13.AddItem "PITTSBURGH GLASS WORKS, LLC"
        .ComboBox13.AddItem "PITTSBURGH GLASS WORKS, ULC"
        .ComboBox13.AddItem "PK USA"
        .ComboBox13.AddItem "PLASCAR INDUSTRIA DE COMPONENTES PLASTICOS LTDA"
        .ComboBox13.AddItem "PLASESS MEXICO, S.A. DE C.V."
        .ComboBox13.AddItem "PLASMAN"
        .ComboBox13.AddItem "PLASTCOAT, A DIVISON OF MAGNA EXTERIORS INC."
        .ComboBox13.AddItem "PLASTECH INC."
        .ComboBox13.AddItem "PLASTIC COMPOUNDERS, INC."
        .ComboBox13.AddItem "PLASTIC OMNIUM ADVANCED INOVATION RESEARCH NV"
        .ComboBox13.AddItem "PLASTIC OMNIUM AUTO INERGY USA LLC"
        .ComboBox13.AddItem "PLASTIC OMNIUM AUTO INERGY ARGENTINA S.A."
        .ComboBox13.AddItem "PLASTIC OMNIUM AUTO INERGY BELGIUM NV"
        .ComboBox13.AddItem "PLASTIC OMNIUM AUTO INERGY SA PTY LTD."
        .ComboBox13.AddItem "PLASTIC OMNIUM AUTOINERGY MEXICO SA DE CV"
        .ComboBox13.AddItem "PLASTIC PLATE INC."
        .ComboBox13.AddItem "PLASTIC RESEARCH & DEVELOPMENT"
        .ComboBox13.AddItem "PLASTIC SERVICE CENTERS"
        .ComboBox13.AddItem "PLASTIC SYSTEMS, LLC."
        .ComboBox13.AddItem "PLASTIC TEC"
        .ComboBox13.AddItem "PLASTIC TRIM INTERNATIONAL INC."
        .ComboBox13.AddItem "PLASTIC-TEC SA DE CV"
        .ComboBox13.AddItem "PLASTICOS ALEDANES"
        .ComboBox13.AddItem "PLASTICOS MAUA LTDA"
        .ComboBox13.AddItem "PLASTICS PLUS INC."
        .ComboBox13.AddItem "PLASTIKON INDUSTRIES, KENTUCKY LLC"
        .ComboBox13.AddItem "PLASTIKON TEXAS, LLC"
        .ComboBox13.AddItem "POLIURETANOS MEXICANOS WOODBRIDGE"
        .ComboBox13.AddItem "POLIX INDUSTRIES"
        .ComboBox13.AddItem "POLYBRITE, A DIVISION OF MAGNA EXTERIORS AND INTERIORS"
        .ComboBox13.AddItem "POLYPLASITCS USA, INC."
        .ComboBox13.AddItem "POLYTEC FOHA"
        .ComboBox13.AddItem "POLYTECH EXCO AUTOMOTIVE"
        .ComboBox13.AddItem "POLYTECH HOLDEN LTD"
        .ComboBox13.AddItem "PORTLAND PRODUCTS, INC."
        .ComboBox13.AddItem "POTENCIA ALLIANCE THAILAND CO., LTD"
        .ComboBox13.AddItem "POWERFLOW INC."
        .ComboBox13.AddItem "PRD INC."
        .ComboBox13.AddItem "PRECISION ASSEMBLIES, INC."
        .ComboBox13.AddItem "PRECISION AUTOMOTIVE PLASTICS"
        .ComboBox13.AddItem "PRECISION PLASTICS"
        .ComboBox13.AddItem "PRECISION POLYMERS"
        .ComboBox13.AddItem "PREMIER SEALS MFG."
        .ComboBox13.AddItem "PRETTY PRODUCTS, INC"
        .ComboBox13.AddItem "PRIME TIME PLASTICS LTD"
        .ComboBox13.AddItem "PRIMERA PLASTICS, INC."
        .ComboBox13.AddItem "PRINCE METAL PRODUCTS"
        .ComboBox13.AddItem "PROACTIVE GROUP"
        .ComboBox13.AddItem "PSI MOLDED PLASTICS"
        .ComboBox13.AddItem "PYA AUTOMOTIVE S DE RL DE CV"
        .ComboBox13.AddItem "PYRAMID PLASTICS"
        .ComboBox13.AddItem "QUALITY CONVERTERS INC."
        .ComboBox13.AddItem "QUALITY MODELS LIMITED"
        .ComboBox13.AddItem "QUALTECH SEATING SYSTEMS"
        .ComboBox13.AddItem "QUANTUM MOLD & ENGINEERING, LLC."
        .ComboBox13.AddItem "QUICKBKS"
        .ComboBox13.AddItem "QUICKSILVER-MFG"
        .ComboBox13.AddItem "RACE MOLD"
        .ComboBox13.AddItem "RAMKO MFG., INC."
        .ComboBox13.AddItem "RANS"
        .ComboBox13.AddItem "RAVAL EUROPE SA"
        .ComboBox13.AddItem "RCO ENGINEERING"
        .ComboBox13.AddItem "RECARO NORTH AMERICA, INC"
        .ComboBox13.AddItem "RECICLADORA JIMESA"
        .ComboBox13.AddItem "RED E PARTS, INC."
        .ComboBox13.AddItem "REHAU INC"
        .ComboBox13.AddItem "REHAU SA DE CV"
        .ComboBox13.AddItem "RESINOID ENGINEERING CORP"
        .ComboBox13.AddItem "REVELATE DESIGNS LLC"
        .ComboBox13.AddItem "REVERE PLASTICS SYSTEMS"
        .ComboBox13.AddItem "REYES AUTOMOTIVE GROUP"
        .ComboBox13.AddItem "RICK YOUNG OUTDOORS LLC"
        .ComboBox13.AddItem "RIDGEVIEW INDUSTRIES"
        .ComboBox13.AddItem "RIETER AUTOMOTIVE CARPET"
        .ComboBox13.AddItem "ROBERT BOSCH LLC AFTERMARKET"
        .ComboBox13.AddItem "ROBERT BOSCH LTDA."
        .ComboBox13.AddItem "ROBIN INDUSTRIES, FREDERICKSBURG FACILITY"
        .ComboBox13.AddItem "ROCKFORD MOLDED PRODUCTS, INC."
        .ComboBox13.AddItem "ROKI AMERICA CO., LTD."
        .ComboBox13.AddItem "ROLLSTAMP MFG."
        .ComboBox13.AddItem "ROUSH MANUFACTURING"
        .ComboBox13.AddItem "ROYAL PLASTICS"
        .ComboBox13.AddItem "ROYAL TECHNOLOGIES CORPORATION"
        .ComboBox13.AddItem "RTI URUGUAY, S.A."
        .ComboBox13.AddItem "RYDER"
        .ComboBox13.AddItem "S & A INDUSTRIES"
        .ComboBox13.AddItem "SA AUTOMOTIVE S. DE R.L. DE C.V."
        .ComboBox13.AddItem "SAARGUMMI TENNESSEE, INC."
        .ComboBox13.AddItem "SAFRAN CABIN CANADA CO."
        .ComboBox13.AddItem "SAGINAW BAY PLASTICS, INC."
        .ComboBox13.AddItem "SAIA-BURGESS AUTOMOTIVE"
        .ComboBox13.AddItem "SAINT GOBAIN BRAZIL"
        .ComboBox13.AddItem "SAINT-GOBAIN MEXICO SA DE CV"
        .ComboBox13.AddItem "SALGA PLASTICS INC."
        .ComboBox13.AddItem "SAMPLES"
        .ComboBox13.AddItem "SAMU DIES CORP."
        .ComboBox13.AddItem "SAN FRANCISCO HAT COMPANY"
        .ComboBox13.AddItem "SANAC PRECISION MEXICO SA DE CV"
        .ComboBox13.AddItem "SANKO GOSEI TECHNOLOGIES USA, INC."
        .ComboBox13.AddItem "SANKO MEXICO SA DE CV"
        .ComboBox13.AddItem "SANKYO AMERICA INC."
        .ComboBox13.AddItem "SANMINA CORPORATION"
        .ComboBox13.AddItem "SANOH AMERICA, INC."
        .ComboBox13.AddItem "SANOH INDUSTRIAL DE MEXICO"
        .ComboBox13.AddItem "SAYASHI INDUSTRY"
        .ComboBox13.AddItem "SCANTRON CORPORATION"
        .ComboBox13.AddItem "SCHOGGI INC. DBA WEST PAW DESIGN"
        .ComboBox13.AddItem "SCOTT INDUSTRIES"
        .ComboBox13.AddItem "SEA LINK INTERNATIONAL"
        .ComboBox13.AddItem "SEASON GROUP USA, LLC."
        .ComboBox13.AddItem "SEATING SYSTEMS OF LAREDO"
        .ComboBox13.AddItem "SEGUE XIAMEN MANUFACTURING SERVICES, INC."
        .ComboBox13.AddItem "SEKISUI KASEI U.S.A., INC."
        .ComboBox13.AddItem "SEMCO PLASTIC COMPANY"
        .ComboBox13.AddItem "SENKO ADVANCED COMPONENTS, INC. DBA SENKO AMERICA"
        .ComboBox13.AddItem "SENSICAL INC"
        .ComboBox13.AddItem "SERENDIPITY ELECTRONICS"
        .ComboBox13.AddItem "SETEX AUTOMOTIVE MEXICO SA DE CV"
        .ComboBox13.AddItem "SETEX TS TECH"
        .ComboBox13.AddItem "SETEX/TST CANADA"
        .ComboBox13.AddItem "SHAMROCK INTERNATIONAL FASTENERS, LLC"
        .ComboBox13.AddItem "SHANGAI AB PLASTIC MOLD CO LTD"
        .ComboBox13.AddItem "SHANGHAI CHINAUST AUTOMOTIVE PLASTICS CORP.,LTD."
        .ComboBox13.AddItem "SHANGHAI DAIMAY AUTOMOTIVE"
        .ComboBox13.AddItem "SHANGHAI NIFCO PLASTIC MANUFACTURER CO., LTD"
        .ComboBox13.AddItem "SHAPE CORP MEXICO"
        .ComboBox13.AddItem "SHAPE CORP."
        .ComboBox13.AddItem "SHARP MFG CO OF AMERICA"
        .ComboBox13.AddItem "SHAYNE ENTERPRISES INC."
        .ComboBox13.AddItem "SHERWOOD INNOVATIONS INC."
        .ComboBox13.AddItem "SHILOH INDUSTRIES, INC."
        .ComboBox13.AddItem "SHINSEI CORPORATION"
        .ComboBox13.AddItem "SHIROKI NORTH AMERICA, INC."
        .ComboBox13.AddItem "SHOETREE DBA DAN'S SHOE REPAIR"
        .ComboBox13.AddItem "SHOUJU INDUSTRIALGROUP LTD.,"
        .ComboBox13.AddItem "SHOWA ALUMINUM CORP. OF AMERICA"
        .ComboBox13.AddItem "SI Plastics"
        .ComboBox13.AddItem "SIEGEL"
        .ComboBox13.AddItem "SIMCO AUTOMOTIVE"
        .ComboBox13.AddItem "SIMMS FISHING PRODUCTS CORP"
        .ComboBox13.AddItem "SISTEMAS MECATRONICOS INTICA SA PI DE CV"
        .ComboBox13.AddItem "SIX MOON DESIGNS"
        .ComboBox13.AddItem "SJ GROUPHK LIMITED"
        .ComboBox13.AddItem "SK TECH, INC."
        .ComboBox13.AddItem "SKULL RACING INDUSTRIA E COMERCIO LTDA"
        .ComboBox13.AddItem "SLE-CO PLASTICS INC."
        .ComboBox13.AddItem "SMALL PARTS INC."
        .ComboBox13.AddItem "SMURFIT-STONE RECYCLING"
        .ComboBox13.AddItem "SOFT ARMOR"
        .ComboBox13.AddItem "SOGEFI AIR & COOLING CANADA CORP."
        .ComboBox13.AddItem "SOGEFI ENGINE"
        .ComboBox13.AddItem "SONOCO PROTECTIVE SOLUTIONS"
        .ComboBox13.AddItem "SONOCO PROTECTIVE SOLUTIONS- FINDLAY"
        .ComboBox13.AddItem "SONY CORPORATION"
        .ComboBox13.AddItem "SOTA TECHNOLOGY INC."
        .ComboBox13.AddItem "SOUTHERN INDIANA PLASTICS"
        .ComboBox13.AddItem "SPECTRA/PREMIUM INDUSTRIES"
        .ComboBox13.AddItem "SPIEWAK & SONS, INC."
        .ComboBox13.AddItem "SPRING HILL SEATING SYSTEMS"
        .ComboBox13.AddItem "SR INJECTION MOLDING"
        .ComboBox13.AddItem "SRG GLOBAL"
        .ComboBox13.AddItem "SRG GLOBAL MEXICO"
        .ComboBox13.AddItem "SRG GLOBAL MEXICO S. DE R.L. DE C.V."
        .ComboBox13.AddItem "STAFFCO DE MEXICO"
        .ComboBox13.AddItem "STANLEY ELECTRIC US CO INC"
        .ComboBox13.AddItem "STANT"
        .ComboBox13.AddItem "STANT MANUFACTURA DE MEXICO S.A DE C.V."
        .ComboBox13.AddItem "STAR MANUFACTURING INC"
        .ComboBox13.AddItem "STAR PLASTICS INC."
        .ComboBox13.AddItem "STEWART INDUSTRIES"
        .ComboBox13.AddItem "STOPOL EQUIPMENT SALES, LLC"
        .ComboBox13.AddItem "STRATFORD PLASTIC COMPONENTS"
        .ComboBox13.AddItem "STRATTEC SECURITY CORPORATION"
        .ComboBox13.AddItem "STRATUS PLASTICS KY LLC"
        .ComboBox13.AddItem "STUDIO D RADIODURANS LLC"
        .ComboBox13.AddItem "SUBARU AUTOMOTIVE"
        .ComboBox13.AddItem "SUBARU SERVICE"
        .ComboBox13.AddItem "SULFO TECHNOLOGIES, LLC"
        .ComboBox13.AddItem "SUMIDA COMPONENTS & MODULES GMBH"
        .ComboBox13.AddItem "SUMIDA COMPONENTS DE MEXICO"
        .ComboBox13.AddItem "SUMIDA SLOVENIJA D.O.O"
        .ComboBox13.AddItem "SUMIRIKO OHIO, INC."
        .ComboBox13.AddItem "SUMIRIKO TENNESSEE, INC."
        .ComboBox13.AddItem "SUMITOMO"
        .ComboBox13.AddItem "SUMMIT PLASTIC MOLDING, INC."
        .ComboBox13.AddItem "SUMMIT PLASTICS SILAO, S DE RL DE CV"
        .ComboBox13.AddItem "SUNDAY AFTERNOONS, INC."
        .ComboBox13.AddItem "SUNFLOWER FASHIONS CO LTD"
        .ComboBox13.AddItem "SUPERIOR PLASTICS"
        .ComboBox13.AddItem "SUPPLY TECHNOLOGIES"
        .ComboBox13.AddItem "SUPPLY TECHNOLOGIES OF CANADA"
        .ComboBox13.AddItem "SUR-FLO PLASTICS & ENG., INC."
        .ComboBox13.AddItem "SYSTEX PRODUCTS CORPORATION"
        .ComboBox13.AddItem "T & M SERVICES"
        .ComboBox13.AddItem "T.A. AMERICA CORPORATION"
        .ComboBox13.AddItem "T.RAD NORTH AMERICA"
        .ComboBox13.AddItem "TAC MANUFACTURING INC"
        .ComboBox13.AddItem "TACHI-S AUTOMOTIVE SEATING U.S.A. LLC"
        .ComboBox13.AddItem "TAESUNG PRECISION CO. LTD"
        .ComboBox13.AddItem "TAG AUTOMOTIVE S.L."
        .ComboBox13.AddItem "TAICA CUBIC PRINTING KY LLC"
        .ComboBox13.AddItem "TAKUMI STAMPING CANADA INC."
        .ComboBox13.AddItem "TAKUMI STAMPING TEXAS, INC."
        .ComboBox13.AddItem "TAKUMI STAMPING, INC."
        .ComboBox13.AddItem "TAMODA APPAREL INC."
        .ComboBox13.AddItem "TARGET PLASTICS TECHNOLOGY"
        .ComboBox13.AddItem "TASUS ALABAMA CORP"
        .ComboBox13.AddItem "TASUS CORPORATION"
        .ComboBox13.AddItem "TASUS TEXAS CORPORATION"
        .ComboBox13.AddItem "TB DE MEXICO SA DE CV"
        .ComboBox13.AddItem "TBDN TENNESSEE"
        .ComboBox13.AddItem "TE CONNECTIVITY CORPORATION"
        .ComboBox13.AddItem "TECH MOLDING SOLUTIONS, LLC."
        .ComboBox13.AddItem "TECHMER PM, LLC"
        .ComboBox13.AddItem "TECSTAR MFG. COMPANY"
        .ComboBox13.AddItem "TEIJIN AUTOMOTIVE TECHNOLOGIES"
        .ComboBox13.AddItem "TEKMART INTEGRATED MANUFACTURING SERVICES"
        .ComboBox13.AddItem "TELAMON"
        .ComboBox13.AddItem "TENNEPLAS"
        .ComboBox13.AddItem "TEPRO- CKR"
        .ComboBox13.AddItem "TEPSO PLASTICS MEX, SA DE C.V"
        .ComboBox13.AddItem "TERNES PROCUREMENT SERVICES"
        .ComboBox13.AddItem "TERNES PROCUREMENT SERVICES."
        .ComboBox13.AddItem "TERRAZIGN INC."
        .ComboBox13.AddItem "TESLA MANUFACTURING BRANDENBURG SE"
        .ComboBox13.AddItem "TESLA MOTORS NETHERLANDS B.V."
        .ComboBox13.AddItem "TETHERTEKS, LLC"
        .ComboBox13.AddItem "TG AUTOMOTIVE SEALING"
        .ComboBox13.AddItem "TG CALIFORNIA AUTO SEALING"
        .ComboBox13.AddItem "TG FLUID SYSTEMS USA CORP"
        .ComboBox13.AddItem "TG KENTUCKY CORP"
        .ComboBox13.AddItem "TG MINTO CORP"
        .ComboBox13.AddItem "TG MISSOURI"
        .ComboBox13.AddItem "TG TEXAS"
        .ComboBox13.AddItem "THAI SUMMIT RAYONG AUTOPARTS INDUSTRY CO., LTD."
        .ComboBox13.AddItem "THB"
        .ComboBox13.AddItem "THE SMELT BELT COMPANY"
        .ComboBox13.AddItem "THE WOODBRIDGE GROUP"
        .ComboBox13.AddItem "THERMAL"
        .ComboBox13.AddItem "THOUGHT TECHNOLOGY LT."
        .ComboBox13.AddItem "THUMB PLASTICS INC."
        .ComboBox13.AddItem "TI AUTOMOTIVE TIANJIN CO., LTD."
        .ComboBox13.AddItem "TI AUTOMOTIVE ARGENTINA S.A."
        .ComboBox13.AddItem "TI AUTOMOTIVE CANADA INC."
        .ComboBox13.AddItem "TI BRASIL"
        .ComboBox13.AddItem "TI FLUID SYSTEMS"
        .ComboBox13.AddItem "TI GROUP AUTOMOTIVE DEESIDE LTD"
        .ComboBox13.AddItem "TI GROUP AUTOMOTIVE S DE R L DE CV"
        .ComboBox13.AddItem "TI GROUP AUTOMOTIVE SYSTEMS S.R.O"
        .ComboBox13.AddItem "TIFCO DONGGUAN CO., LTD."
        .ComboBox13.AddItem "TIGERPOLY INDUSTRIA DE MEXICO S.A. DE C.V."
        .ComboBox13.AddItem "TIMBUK2"
        .ComboBox13.AddItem "TK HOLDINGS INC."
        .ComboBox13.AddItem "TMD WEK NORTH LLC"
        .ComboBox13.AddItem "TMD WEK SOUTH LLC"
        .ComboBox13.AddItem "TMM BAJA CALIFORNIA, S. DE R.L DE C.V."
        .ComboBox13.AddItem "TOA E&I AMERICA, INC."
        .ComboBox13.AddItem "TOM BIHN, INC."
        .ComboBox13.AddItem "TOM SMITH INDUSTRIES"
        .ComboBox13.AddItem "TOMASCO MULCIBER"
        .ComboBox13.AddItem "TOOLING VENTURES INC."
        .ComboBox13.AddItem "TOPRE AMERICA CORPORATION"
        .ComboBox13.AddItem "TOTAL NETWORK MANUFACTURING, LLC"
        .ComboBox13.AddItem "TOWER INTERNATIONAL, INC"
        .ComboBox13.AddItem "TOYO AUTOMOTIVE PARTS USA, INC."
        .ComboBox13.AddItem "TOYO SEAT USA CORP"
        .ComboBox13.AddItem "TOYODA GOSEI AUTOMOTIVE SEALING MEXICO"
        .ComboBox13.AddItem "TOYODA GOSEI AUTOMOTIVE SEALING MEXICO SA DE CV"
        .ComboBox13.AddItem "TOYODA GOSEI IRAPUATO MEXICO S.A.DE C.V."
        .ComboBox13.AddItem "TOYODA GOSEI IRAPUATO MEXICO,S.A. DE C.V."
        .ComboBox13.AddItem "TOYODABO MFG KENTUCKY"
        .ComboBox13.AddItem "Toyota Baja California"
        .ComboBox13.AddItem "Toyota Boshoku America"
        .ComboBox13.AddItem "TOYOTA BOSHOKU AMERICA."
        .ComboBox13.AddItem "TOYOTA BOSHOKU CANADA, INC."
        .ComboBox13.AddItem "TOYOTA BOSHOKU INDIANA"
        .ComboBox13.AddItem "Toyota Canada"
        .ComboBox13.AddItem "Toyota Indiana"
        .ComboBox13.AddItem "TOYOTA KENTUCKY"
        .ComboBox13.AddItem "TOYOTA LOGISTICS SERVICES"
        .ComboBox13.AddItem "TOYOTA MOTOR ENGINEERING & MANUFACTURING NORTH AMERICA, INC"
        .ComboBox13.AddItem "TOYOTA MOTOR MANUFACTURING"
        .ComboBox13.AddItem "TOYOTA MOTOR MANUFACTURING TMMGT"
        .ComboBox13.AddItem "TOYOTA MOTOR MANUFACTURING DE GUANAJUATO, S.A. DE C.V."
        .ComboBox13.AddItem "TOYOTA MOTOR NORTH AMERICA, INC."
        .ComboBox13.AddItem "Toyota Northern Kentucky SIA"
        .ComboBox13.AddItem "Toyota West Virginia"
        .ComboBox13.AddItem "Toyota Service Midwest"
        .ComboBox13.AddItem "Toyota Texas"
        .ComboBox13.AddItem "TOYOTA TSUSHO AMERICA"
        .ComboBox13.AddItem "TOYOTA TSUSHO CANADA"
        .ComboBox13.AddItem "TOYOTETSU AMERICA, INC."
        .ComboBox13.AddItem "TOYOTETSU CANADA, INC."
        .ComboBox13.AddItem "TOYOTETSU MID-AMERICA INC."
        .ComboBox13.AddItem "TRANSISTOR DEVICES, INC."
        .ComboBox13.AddItem "TRANSNAV TECHNOLOGIES"
        .ComboBox13.AddItem "TRANSNAV TECHNOLOGIES INC."
        .ComboBox13.AddItem "TRANSTRON AMERICA INC."
        .ComboBox13.AddItem "TRI-CON INDUSTRIES"
        .ComboBox13.AddItem "TRIM MASTERS"
        .ComboBox13.AddItem "TRIMOLD LLC"
        .ComboBox13.AddItem "TRI-PARAGON"
        .ComboBox13.AddItem "TRIN, INC."
        .ComboBox13.AddItem "TRIPAC INTERNATIONAL, INC."
        .ComboBox13.AddItem "TRQSS"
        .ComboBox13.AddItem "TRUE NORTH"
        .ComboBox13.AddItem "TRULIFE"
        .ComboBox13.AddItem "TRW AUTOMOTIVE-OSS"
        .ComboBox13.AddItem "TS TECH ALABAMA"
        .ComboBox13.AddItem "TS TECH CANADA"
        .ComboBox13.AddItem "TS TECH DEUTSCHLAND GMBH"
        .ComboBox13.AddItem "TS TECH INDIANA LLC"
        .ComboBox13.AddItem "TS TECH NORTH AMERICA"
        .ComboBox13.AddItem "TS TECH USA"
        .ComboBox13.AddItem "TST NA TRIM LLC"
        .ComboBox13.AddItem "TURTLE FUR COMPANY"
        .ComboBox13.AddItem "TYLER MANUFACTURING COMPANY"
        .ComboBox13.AddItem "U.S. FARATHANE, S.A. DE C.V."
        .ComboBox13.AddItem "UDDER TECH, INC."
        .ComboBox13.AddItem "UGN"
        .ComboBox13.AddItem "UGN INC."
        .ComboBox13.AddItem "ULTRA MANUFACTURING LIMITED"
        .ComboBox13.AddItem "ULTRA MANUFACTURING SA DE CV"
        .ComboBox13.AddItem "UNIFORM COLOR COMPANY"
        .ComboBox13.AddItem "UNION NIFCO"
        .ComboBox13.AddItem "UNIPRES MEXICANA"
        .ComboBox13.AddItem "UNIPRES MEXICANA, S.A. DE C.V."
        .ComboBox13.AddItem "UNIQUE ASSEMBLY & DECORATING"
        .ComboBox13.AddItem "UNIQUE FABRICATING DE MEXICO, S.A. DE C.V."
        .ComboBox13.AddItem "UNIQUE FABRICATING, INC."
        .ComboBox13.AddItem "US FARATHANE CORPORATION"
        .ComboBox13.AddItem "US YACHIYO"
        .ComboBox13.AddItem "UV RAVEN, LLC"
        .ComboBox13.AddItem "VAL PRODUCTS INC."
        .ComboBox13.AddItem "VALEO NORTH AMERICA, INC."
        .ComboBox13.AddItem "VALLEN DISTRIBUTION"
        .ComboBox13.AddItem "VALLEY ENTERPRISES"
        .ComboBox13.AddItem "VENTRA EVART, LLC"
        .ComboBox13.AddItem "VENTRA FOWLERVILLE"
        .ComboBox13.AddItem "VENTRA GRAND RAPIDS, LLC"
        .ComboBox13.AddItem "VENTRA IONIA LLC"
        .ComboBox13.AddItem "VENTRA PLASTICS"
        .ComboBox13.AddItem "VENTRA PLASTICS KITCHENER"
        .ComboBox13.AddItem "VENTRA PLASTICS PETERBOROUGH"
        .ComboBox13.AddItem "VENTRA SALEM, LLC"
        .ComboBox13.AddItem "VFM, LLC."
        .ComboBox13.AddItem "VIAM MFG"
        .ComboBox13.AddItem "VIDON PLASTICS, INC."
        .ComboBox13.AddItem "VIDRIO"
        .ComboBox13.AddItem "VINCENT INDUSTRIAL"
        .ComboBox13.AddItem "VINTEC COMPANY"
        .ComboBox13.AddItem "VINTECH INDUSTRIES"
        .ComboBox13.AddItem "VISTA INDUSTRIAL PACKAGING LLC"
        .ComboBox13.AddItem "VISTECH"
        .ComboBox13.AddItem "VISTEON"
        .ComboBox13.AddItem "VITESCO TECHNOLOGIES ROMANIA SRL"
        .ComboBox13.AddItem "VITESCO TECHNOLOGIES USA, LLC"
        .ComboBox13.AddItem "VITRO AUTOMOTRIZ, S.A. DE C.V."
        .ComboBox13.AddItem "VNO DESIGN & ENGINEERING"
        .ComboBox13.AddItem "VOESTALPINE ROTEC SUMMO CORP."
        .ComboBox13.AddItem "VOGTRONICS GMBH"
        .ComboBox13.AddItem "VOLKSWAGEN DE MEXICO SA DE CV"
        .ComboBox13.AddItem "VOLKSWAGEN GROUP OF AMERICA SERVICE"
        .ComboBox13.AddItem "VOLKSWAGEN OF AMERICA"
        .ComboBox13.AddItem "VOSS AUTOMOTIVE, INC."
        .ComboBox13.AddItem "VPI ACQUISITION LLC, DBA VIKING PLASTICS"
        .ComboBox13.AddItem "VR MANUFACTURING MEXICO, S. DE R.L. DE C.V."
        .ComboBox13.AddItem "VU MANUFACTURING"
        .ComboBox13.AddItem "VUTEQ CANADA INC"
        .ComboBox13.AddItem "VUTEQ GUANAJUATO MEXICO, S.A. DE C.V."
        .ComboBox13.AddItem "VUTEQ SERVICE MEXICO SA DE CV"
        .ComboBox13.AddItem "VUTEQ USA, INC."
        .ComboBox13.AddItem "VUTEX INC."
        .ComboBox13.AddItem "W & E SALES CO., INC."
        .ComboBox13.AddItem "W.E.T. AUTOMOTIVE SYSTEMS LTD."
        .ComboBox13.AddItem "W.K. INDUSTRIES, INC."
        .ComboBox13.AddItem "WABASH PLASTICS, INC."
        .ComboBox13.AddItem "WALKAPOCKET LLC"
        .ComboBox13.AddItem "WALLACH IRON & METAL CO INC."
        .ComboBox13.AddItem "WARN INDUSTRIES"
        .ComboBox13.AddItem "WATERVILLE TG"
        .ComboBox13.AddItem "WC&R INTERESTS, LLC DBA DIAMOND BRAND"
        .ComboBox13.AddItem "WEST MICHIGAN FLOCKING"
        .ComboBox13.AddItem "WEST MICHIGAN MOLDING, INC."
        .ComboBox13.AddItem "WEST TROY LLC"
        .ComboBox13.AddItem "WESTCOMB OUTERWEAR"
        .ComboBox13.AddItem "WESTERN MOUNTAINEERING"
        .ComboBox13.AddItem "WIA MOLDING LLC"
        .ComboBox13.AddItem "WILD THINGS LLC"
        .ComboBox13.AddItem "WILLIAMSBURG MFG."
        .ComboBox13.AddItem "WINDSOR MACHINE & STAMPING LTD"
        .ComboBox13.AddItem "WINDSOR MOLD SALINE"
        .ComboBox13.AddItem "WINNER SPORTSWEAR LTD"
        .ComboBox13.AddItem "WINTERGREEN NORTHERN WEAR"
        .ComboBox13.AddItem "WITTE NEJDEK SPOL. S R.O."
        .ComboBox13.AddItem "WKW NORTH AMERICA, LLC"
        .ComboBox13.AddItem "WOODBRIDGE FOAM CORPORATION"
        .ComboBox13.AddItem "WOODBRIDGE FOAM CORPORATION-KANSAS CITY"
        .ComboBox13.AddItem "WOOKEY DESIGN STUDIO"
        .ComboBox13.AddItem "WORLD CLASS INDUSTRIES IL INC."
        .ComboBox13.AddItem "World Resource Solution"
        .ComboBox13.AddItem "WREX INDUSTRIES"
        .ComboBox13.AddItem "WURTH DMB SUPPLY"
        .ComboBox13.AddItem "WURTH INDUSTRIAL US, INC."
        .ComboBox13.AddItem "XEMAR"
        .ComboBox13.AddItem "XIN POINT MEXICO, S. DE R.L. DE C.V."
        .ComboBox13.AddItem "XINQUAN MEXICO AUTOMOTIVE TRIM S DE RL DE CV."
        .ComboBox13.AddItem "YACHIO - HONDA OF ALABAMA"
        .ComboBox13.AddItem "YACHIYO MANUFACTURING OF AMERICA"
        .ComboBox13.AddItem "YACHIYO MEXICO MANUFACTURING SA DE CV"
        .ComboBox13.AddItem "YACHIYO MEXICO MFG SA DE CV"
        .ComboBox13.AddItem "YACHIYO OF AMERICA INC."
        .ComboBox13.AddItem "YAHAGI AMERICA MOLDING, INC."
        .ComboBox13.AddItem "YANFENG"
        .ComboBox13.AddItem "YANFENG MEXICO INTERIORS S.DE.R.L.C.V."
        .ComboBox13.AddItem "YANFENG US AUTOMOTIVE INTERIOR SYSTEMS I, LLC"
        .ComboBox13.AddItem "YANFENG US AUTOMOTIVE INTERIOR SYSTEMS II, LLC"
        .ComboBox13.AddItem "YAPP INDIA AUTOMOTIVE SYSTEMS PVT LTD"
        .ComboBox13.AddItem "YAPP USA AUTOMOTIVE SYSTEMS INC."
        .ComboBox13.AddItem "YASUFUKU USA, INC."
        .ComboBox13.AddItem "YAZAKI CIEMEL S.A."
        .ComboBox13.AddItem "YAZAKI NASHVILLE"
        .ComboBox13.AddItem "YAZAKI NORTH AMERICA, INC."
        .ComboBox13.AddItem "YAZAKI NORTH AMERICA INC - TX"
        .ComboBox13.AddItem "YOKOHAMA INDUSTRIES AMERICAS"
        .ComboBox13.AddItem "YUSA AUTOPARTS MEXICO, S.A. DE C.V."
        .ComboBox13.AddItem "YUSA"
        .ComboBox13.AddItem "ZEAGLE SYSTEMS, INC."
        .ComboBox13.AddItem "ZF WASHINGTON A435"
        .ComboBox13.AddItem "ZIMMERMANN-TECHNIK HONG KONG LTD."

        
    
    End With

End Sub

Sub MoldFormReset()
' this load the molded section for add item on kickoff sheet

    With MoldForm
    

    .ComboBox2.Clear
    .ComboBox2.AddItem "31-60"
    .ComboBox2.AddItem "61-90"
    .ComboBox2.AddItem "91-110"
    .ComboBox2.AddItem "111-130"
    .ComboBox2.AddItem "131-180"
    .ComboBox2.AddItem "181-230"
    .ComboBox2.AddItem "231-280"
    .ComboBox2.AddItem "281-360"
    .ComboBox2.AddItem "361-460"
    .ComboBox2.AddItem "461"
    
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
    .ComboBox3.AddItem "5"
    
    .ComboBox4.Clear
    .ComboBox4.AddItem "0"
    .ComboBox4.AddItem "1"
    .ComboBox4.AddItem "2"
    .ComboBox4.AddItem "3"
    .ComboBox4.AddItem "4"
    .ComboBox4.AddItem "5"
    
    End With
    
    
End Sub
Sub CompFormReset()

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
Sub submitBom()
'this takes info from userform and loads it into Kickoff sheet

Set opRng = Worksheets("Kickoff Boms").Range("B9:B35")

iRow = Cells(Rows.Count, 2).End(xlUp).Row + 1

    Dim ph As Worksheet
    
    Set ph = ThisWorkbook.Sheets("Kickoff Boms")
 
    
    With ph
        
        
        
        .Cells(iRow, 1) = Cells(iRow, 1).Offset(-1, 0) + 1
        .Cells(iRow, 2) = BomForm.TextBox1.Value
        .Cells(iRow, 3) = BomForm.TextBox2.Value
        .Cells(iRow, 4) = BomForm.ComboBox1.Value
        .Cells(iRow, 5) = BomForm.TextBox4.Value
        .Cells(iRow, 6) = BomForm.ComboBox10.Value
        .Cells(iRow, 7) = BomForm.ComboBox2.Value
        .Cells(iRow, 8) = BomForm.ComboBox3.Value
        .Cells(iRow, 9) = BomForm.TextBox5.Value
        .Cells(iRow, 10) = BomForm.ComboBox4.Value
        .Cells(iRow, 11) = BomForm.TextBox6.Value
        .Cells(iRow, 12) = BomForm.TextBox7.Value
        .Cells(iRow, 13) = BomForm.TextBox6.Value
        .Cells(iRow, 15) = BomForm.TextBox8.Value
        .Cells(iRow, 30) = BomForm.ComboBox11.Value
        .Cells(iRow, 35) = BomForm.ComboBox12.Value
        If BomForm.CheckBox1.Value = True Then Cells(iRow, 36) = "MEX"
        
    End With
        
        
End Sub

Sub submitMold()
'take info from mold form and loads it into Kickoff sheet
    Dim xRng As Range, xCell As Range
    Dim i As Integer
    Dim th As Worksheet
    Dim text1 As String
    Dim text2 As String
    
    Set th = ThisWorkbook.Sheets("Kickoff Boms")

    
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
        .Cells(iRow, 16) = MoldForm.ComboBox2.Value
        .Cells(iRow, 17) = MoldForm.ComboBox1.Value
        .Cells(iRow, 31) = MoldForm.ComboBox3.Value
        .Cells(iRow, 32) = MoldForm.ComboBox4.Value
        .Cells(iRow, 33) = MoldForm.TextBox7.Value
        

        
    End With
        
        
End Sub
Sub submitPackaging()
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
    Set LocRng = ThisWorkbook.Worksheets("LocArray").Range("v3:v7")
    
    
    Set ph = ThisWorkbook.Sheets("Kickoff Boms")

    ThisWorkbook.Sheets("LocArray").Cells(3, 22) = Packaging.TextBox1.Value
    ThisWorkbook.Sheets("LocArray").Cells(4, 22) = Packaging.TextBox2.Value
    ThisWorkbook.Sheets("LocArray").Cells(5, 22) = Packaging.TextBox3.Value
    ThisWorkbook.Sheets("LocArray").Cells(6, 22) = Packaging.TextBox4.Value
    ThisWorkbook.Sheets("LocArray").Cells(7, 22) = Packaging.TextBox5.Value
    
    
    ThisWorkbook.Sheets("LocArray").Cells(3, 23) = Packaging.TextBox6.Value
    ThisWorkbook.Sheets("LocArray").Cells(4, 23) = Packaging.TextBox7.Value
    ThisWorkbook.Sheets("LocArray").Cells(5, 23) = Packaging.TextBox8.Value
    ThisWorkbook.Sheets("LocArray").Cells(6, 23) = Packaging.TextBox9.Value
    ThisWorkbook.Sheets("LocArray").Cells(7, 23) = Packaging.TextBox10.Value
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

Sub submitComp()
' this takes info from comp form and load it into kickoff sheet
ThisWorkbook.Worksheets("LocArray").Unprotect Password:="1234"
 Dim ph As Worksheet
    Dim text1 As String
    Dim text2 As String
    Dim text3 As String
    Dim text4 As String
    Dim rng As Range
    Dim LocRng As Range
    Set LocRng = ThisWorkbook.Worksheets("LocArray").Range("q3:q12")
    
    
    Set ph = ThisWorkbook.Sheets("Kickoff Boms")

    If CompForm.CheckBox1.Value = True Then Cells(iRow, 34) = "Yes"
    

    ThisWorkbook.Sheets("LocArray").Cells(3, 17) = CompForm.TextBox1.Value
    ThisWorkbook.Sheets("LocArray").Cells(4, 17) = CompForm.TextBox2.Value
    ThisWorkbook.Sheets("LocArray").Cells(5, 17) = CompForm.TextBox3.Value
    ThisWorkbook.Sheets("LocArray").Cells(6, 17) = CompForm.TextBox4.Value
    ThisWorkbook.Sheets("LocArray").Cells(7, 17) = CompForm.TextBox5.Value
    ThisWorkbook.Sheets("LocArray").Cells(8, 17) = CompForm.TextBox6.Value
    ThisWorkbook.Sheets("LocArray").Cells(9, 17) = CompForm.TextBox7.Value
    ThisWorkbook.Sheets("LocArray").Cells(10, 17) = CompForm.TextBox8.Value
    ThisWorkbook.Sheets("LocArray").Cells(11, 17) = CompForm.TextBox9.Value
    ThisWorkbook.Sheets("LocArray").Cells(12, 17) = CompForm.TextBox10.Value
    
    
    ThisWorkbook.Sheets("LocArray").Cells(3, 18) = CompForm.TextBox31.Value
    ThisWorkbook.Sheets("LocArray").Cells(4, 18) = CompForm.TextBox32.Value
    ThisWorkbook.Sheets("LocArray").Cells(5, 18) = CompForm.TextBox33.Value
    ThisWorkbook.Sheets("LocArray").Cells(6, 18) = CompForm.TextBox34.Value
    ThisWorkbook.Sheets("LocArray").Cells(7, 18) = CompForm.TextBox35.Value
    ThisWorkbook.Sheets("LocArray").Cells(8, 18) = CompForm.TextBox36.Value
    ThisWorkbook.Sheets("LocArray").Cells(9, 18) = CompForm.TextBox37.Value
    ThisWorkbook.Sheets("LocArray").Cells(10, 18) = CompForm.TextBox38.Value
    ThisWorkbook.Sheets("LocArray").Cells(11, 18) = CompForm.TextBox39.Value
    ThisWorkbook.Sheets("LocArray").Cells(12, 18) = CompForm.TextBox40.Value
    
    
    ThisWorkbook.Sheets("LocArray").Cells(3, 19) = CompForm.TextBox41.Value
    ThisWorkbook.Sheets("LocArray").Cells(4, 19) = CompForm.TextBox42.Value
    ThisWorkbook.Sheets("LocArray").Cells(5, 19) = CompForm.TextBox43.Value
    ThisWorkbook.Sheets("LocArray").Cells(6, 19) = CompForm.TextBox44.Value
    ThisWorkbook.Sheets("LocArray").Cells(7, 19) = CompForm.TextBox45.Value
    ThisWorkbook.Sheets("LocArray").Cells(8, 19) = CompForm.TextBox46.Value
    ThisWorkbook.Sheets("LocArray").Cells(9, 19) = CompForm.TextBox47.Value
    ThisWorkbook.Sheets("LocArray").Cells(10, 19) = CompForm.TextBox48.Value
    ThisWorkbook.Sheets("LocArray").Cells(11, 19) = CompForm.TextBox49.Value
    ThisWorkbook.Sheets("LocArray").Cells(12, 19) = CompForm.TextBox50.Value
    
    
    ThisWorkbook.Sheets("LocArray").Cells(3, 20) = CompForm.ComboBox1.Value
    ThisWorkbook.Sheets("LocArray").Cells(4, 20) = CompForm.ComboBox2.Value
    ThisWorkbook.Sheets("LocArray").Cells(5, 20) = CompForm.ComboBox3.Value
    ThisWorkbook.Sheets("LocArray").Cells(6, 20) = CompForm.ComboBox4.Value
    ThisWorkbook.Sheets("LocArray").Cells(7, 20) = CompForm.ComboBox5.Value
    ThisWorkbook.Sheets("LocArray").Cells(8, 20) = CompForm.ComboBox6.Value
    ThisWorkbook.Sheets("LocArray").Cells(9, 20) = CompForm.ComboBox7.Value
    ThisWorkbook.Sheets("LocArray").Cells(10, 20) = CompForm.ComboBox8.Value
    ThisWorkbook.Sheets("LocArray").Cells(11, 20) = CompForm.ComboBox9.Value
    ThisWorkbook.Sheets("LocArray").Cells(12, 20) = CompForm.ComboBox10.Value

    
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

Sub ShowBom_Form()
'call for BOM form
    BomForm.Show
End Sub
Sub ShowMold_Form()
'call for Mold form
    MoldForm.Show
End Sub
Sub ShowComp_Form()
'call for Comp form
    CompForm.Show
End Sub

Sub enterAssembly()
'sub for entering cost elements, based on item type, assembly

slow2
Application.SendKeys ("Material Overhead")
Application.SendKeys ("{Tab}")
Application.SendKeys ("Yield-PRJ")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
slow1
Application.SendKeys ".01"
slow1

Application.SendKeys ("{down}")
Application.SendKeys ("Material Overhead")
Application.SendKeys ("{Tab}")
Application.SendKeys ("Admin1-PRJ")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
slow1
Application.SendKeys ".06"
slow1

Application.SendKeys ("{down}")
Application.SendKeys ("Material Overhead")
Application.SendKeys ("{Tab}")
Application.SendKeys ("TchCoASMP")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
slow1
Application.SendKeys ".04"
slow1

Application.SendKeys ("{down}")
Application.SendKeys ("Material Overhead")
Application.SendKeys ("{Tab}")
Application.SendKeys ("Rylty-PRJ")
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
    
End Sub
Sub enterShootShip()
'sub for entering cost elements, based on item type, shoot and ship

GrabToolCost

slow2
Application.SendKeys ("Material Overhead")
Application.SendKeys ("{Tab}")
Application.SendKeys ("Yield-PRJ")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
slow1
Application.SendKeys ".01"
slow1

Application.SendKeys ("{down}")
Application.SendKeys ("Material Overhead")
Application.SendKeys ("{Tab}")
Application.SendKeys ("Aux Co-PRJ")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
slow1
Application.SendKeys ".14"
slow1

Application.SendKeys ("{down}")
Application.SendKeys ("Material Overhead")
Application.SendKeys ("{Tab}")
Application.SendKeys ("TechCoMLDP")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
slow1
Application.SendKeys ".06"
slow1

Application.SendKeys ("{down}")
Application.SendKeys ("Material Overhead")
Application.SendKeys ("{Tab}")
Application.SendKeys ("Tool RprP")
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
Application.SendKeys ("Rylty-PRJ")
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
    Application.SendKeys ("Glass1-PRJ")
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
Application.SendKeys ("Yield-PRJ")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
slow1
Application.SendKeys ".01"
slow1

Application.SendKeys ("{down}")
Application.SendKeys ("Material Overhead")
Application.SendKeys ("{Tab}")
Application.SendKeys ("Aux Co-PRJ")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
slow1
Application.SendKeys ".14"
slow1

Application.SendKeys ("{down}")
Application.SendKeys ("Material Overhead")
Application.SendKeys ("{Tab}")
Application.SendKeys ("TechCoMLDP")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
slow1
Application.SendKeys ".06"
slow1

Application.SendKeys ("{down}")
Application.SendKeys ("Material Overhead")
Application.SendKeys ("{Tab}")
Application.SendKeys ("Tool RprP")
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
    Application.SendKeys ("Glass1-PRJ")
    Application.SendKeys ("{Tab}")
    Application.SendKeys ("{Tab}")
    Application.SendKeys ("{Tab}")
    slow1
    Application.SendKeys Cells(findCell.Row, 13).Value
    slow1
End If
    
End Sub
Sub BringToFront()
'brings kickoff boms sheet into focus
    Dim setFocus As Long
    
    ThisWorkbook.Worksheets("Kickoff Boms").Activate
    setFocus = SetForegroundWindow(Application.hwnd)
End Sub

Sub EndBOM()
'gives error msg and ends process
BringToFront
MsgBox "Out of Alignment ending Bom Entry"
DoEvents
End


End Sub

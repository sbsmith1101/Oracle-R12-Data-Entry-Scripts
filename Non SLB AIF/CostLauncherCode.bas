Attribute VB_Name = "CostLauncherCode"
Public Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As LongPtr, ByVal y As LongPtr) As LongPtr
Public Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwmilliseconds As LongPtr)

Public Declare PtrSafe Sub mouse_event Lib "user32" (ByVal dwFlags As LongPtr, ByVal dx As Long, ByVal dy As LongPtr, ByVal cButtons As LongPtr, ByVal swextrainfo As LongPtr) '
Public Const mouseeventf_Leftdown = &H2
Public Const mouseeventf_Leftup = &H4
Public Const mouseeventf_Rightdown As Long = &H8
Public Const mouseeventf_rightup As Long = &H10

Option Explicit
Public txt1 As String
Public Item As String
Public OrgCodeStart As String
Public orgCode As String
Public OrgNum As String
Public CostType As String
Public AcctCode As String
Public MorA As String
Public findCell As Range
Public AifType As String
Public cycleLoop As Integer
Public t As Integer
Public ExitCon As Boolean
Public retry1 As Boolean


Public Sub ClickOnCornerWindow()
' this sub clicks on top left of primary screen, using this to click on window to focably bring it into focus.
Dim oLeft As Long
oLeft = 100  'last desktop size
oLeft = 80   'revised desktop

Dim OTop As Long
OTop = 100

SetCursorPos oLeft, OTop

mouse_event mouseeventf_Leftdown, 0, 0, 0, 0
mouse_event mouseeventf_Leftup, 0, 0, 0, 0


End Sub
Sub ActivateOracle()
' WIP
Dim MyAppID, ReturnValue
AppActivate "Oracle Applications", True
'need to bring to front




End Sub
Sub CopyCompareCell()
'this sub activates a crtl C action and then pulls the info in the windows clipboard and stores it in a variable so it can be evaluated at various intervals
Dim ClipObj As New DataObject
ClipObj.SetText Text:=Empty
ClipObj.PutInClipboard
Application.SendKeys ("^c")
slow1



On Error Resume Next
ClipObj.GetFromClipboard
On Error GoTo 0
On Error Resume Next
    txt1 = ClipObj.GetText(1)
On Error GoTo 0

End Sub
Sub ClearClipboard()
'this does not appear to work, check with msgbox
ThisWorkbook.Worksheets("LocArray").Cells(a, 2).Copy

Stop
End Sub
Sub openMyFile()
'this sub is for taking a item # and searching the submissions dirctory to open associated folder.
Dim rng As Range
Set rng = ThisWorkbook.Worksheets("AIF").Range("B4:B45")
Set findCell = rng.Find(what:="*", searchFormat:=True)
    If (findCell Is Nothing) Then
        BringToFront
        MsgBox "All Reports Pulled"
        Exit Sub
    End If

Dim MainPath As String
Dim subSum As String
Dim FullPath As String

MainPath = "\\srv-corp-nas01\SUBMISSIONS\"
subSum = Left(findCell, 3) & "00"
'subSum = Left(ActiveCell.Value, 3) & "00"
'FullPath = MainPath & subSum & "\" & Item
Dim a As String
a = Dir(MainPath & subSum & "\" & findCell.Value & "*", vbDirectory)


FullPath = MainPath & subSum & "\" & a


Call Shell("explorer.exe" & " " & FullPath, vbNormalFocus)
On Error GoTo Label

GoTo FINISH:

Label:
ThisWorkbook.Worksheets("AIF").Cells(findCell.Row, 1) = "Folder Error"
ThisWorkbook.Worksheets("AIF").Cells(findCell.Row, 1).Interior.ColorIndex = 7

FINISH:
End Sub
Sub checkFileName()

MsgBox Dir("\\srv-corp-nas01\SUBMISSIONS\28400\28412*", vbDirectory)

End Sub
Sub opensingleFile()
' single file version of previous sub
Dim rng As Range


Dim MainPath As String
Dim subSum As String
Dim FullPath As String


MainPath = "\\srv-corp-nas01\SUBMISSIONS\"
subSum = Left(ActiveCell.Value, 3) & "00"

'FullPath = MainPath & subSum & "\" & a


Dim a As String
a = Dir(MainPath & subSum & "\" & ActiveCell.Value & "*", vbDirectory)


FullPath = MainPath & subSum & "\" & a
Call Shell("explorer.exe" & " " & FullPath, vbNormalFocus)


End Sub
Sub pull1Report()
'sequence of report subs for calling cost sheet and supply chain reports from oracle, calls value of highlighted cell
Set findCell = ActiveCell
DoEvents
'ClickOnCornerWindow
Application.SendKeys ("%tt"), True
slow1
pullCostReports
slow2
BringToFront
End Sub

Sub slow1()
'sub to cause 1 sec delay in process
DoEvents
Application.Wait (Now + TimeValue("00:00:01"))

End Sub

Sub slow2()
'sub to cause 2 sec delay in process
DoEvents
Application.Wait (Now + TimeValue("00:00:02"))
End Sub
Sub pullCostReports()
'sequence of report subs for calling cost sheet and supply chain reports from oracle, item # supplied by variable
ExitCon = False

setvars
UpdateAcctCode
slow2
If ExitCon = True Then GoTo SkipToEnd
slow2
Application.SendKeys ("%tt"), True
slow1
DoEvents

changeOrg
slow2

slow1
Application.SendKeys ("3"), True

slow1
Application.SendKeys (orgCode), True

slow2
Application.SendKeys ("~"), True
slow2
Application.SendKeys ("Supply Chain Cost Rollup - Print Report"), True
slow1
Application.SendKeys ("{Tab}"), True

Application.SendKeys ("{Tab}"), True

Application.SendKeys (CostType), True
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
slow2
CopyCompareCell
If Not txt1 = "Corporate" Then BringToFront
If Not txt1 = "Corporate" Then MsgBox ("Out of alignment, sub exited")
If Not txt1 = "Corporate" Then ExitCon = True
If Not txt1 = "Corporate" Then GoTo SkipToEnd


Application.SendKeys ("{Tab}")
Application.SendKeys ("Single level rollup")

Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
slow1
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}"), True
slow1
Application.SendKeys ("{Tab}"), True
Application.SendKeys ("{Tab}")
slow1
Application.SendKeys ("{Tab}"), True
Application.SendKeys ("{Tab}")

Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
slow1
Application.SendKeys ("{Tab}"), True
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}"), True

Application.SendKeys ("{Tab}")
slow1
Application.SendKeys (Item)
slow1
Application.SendKeys ("{Tab}")
Application.SendKeys ("~"), True
slow1
Application.SendKeys ("~"), True
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
Application.SendKeys ("~"), True
slow1
Application.SendKeys ("{Tab}")
Application.SendKeys ("~"), True
slow2


slow1
Application.SendKeys ("%tt"), True

DoEvents
Application.Wait (Now + TimeValue("00:00:03"))
Application.SendKeys ("7"), True
slow1
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
slow1
Application.SendKeys ("~"), True
slow1

Application.SendKeys ("{Tab}")
Application.SendKeys ("~"), True
slow1
Application.SendKeys ("Custom Item Cost Report"), True
Application.SendKeys ("{Tab}"), True
slow1
Application.SendKeys (CostType)
Application.SendKeys ("{Tab}"), True
Application.SendKeys (Item)
Application.SendKeys ("{Tab}")
slow2
Application.SendKeys ("~"), True
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
slow2
CopyCompareCell

If Not txt1 = "No" Then BringToFront
If Not txt1 = "No" Then MsgBox ("Out of alignment, sub exited")
If Not txt1 = "No" Then ExitCon = True
If Not txt1 = "No" Then GoTo SkipToEnd
slow1
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
slow1
Application.SendKeys ("~"), True
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
slow1
Application.SendKeys ("~"), True
slow1
Application.SendKeys ("{Tab}")
Application.SendKeys ("~"), True


slow1
Application.SendKeys ("{F4}")

openMyFile

slow2
SkipToEnd:
Application.SendKeys ("{NUMLOCK}"), True
End Sub
Sub ChkInhouseCost()


If findCell.Offset(0, 9) = "0" Then GoTo EmptSkip
If findCell.Offset(0, 9) = "" Then GoTo EmptSkip
ClickOnCornerWindow
txt1 = 0
slow2
slow2
slow2
Application.SendKeys ("%tt"), True
slow1
Application.SendKeys ("1"), True
slow1
Application.SendKeys Item

Application.SendKeys ("{Tab}")
slow2
'Application.SendKeys ("Frozen"), True
Application.SendKeys CostType
Application.SendKeys ("%i"), True
slow1
Application.SendKeys ("{Tab}")

CopyCompareCell

'If Not txt1 = "Frozen" Then
If Not LCase(txt1) = LCase(CostType) Then
    BringToFront
    MsgBox "Out of place Error, please check cost for item" & Item
    End
End If
Application.SendKeys ("%i"), True
slow1
Application.SendKeys ("Total cos"), True
slow1
CopyCompareCell
On Error Resume Next
If txt1 > findCell.Offset(0, 9) Then findCell.Interior.ColorIndex = 9
If txt1 = 0 Then findCell.Offset(0, 9).Interior.ColorIndex = 12
On Error GoTo 0
slow2
Application.SendKeys ("%fc"), True



EmptSkip:

End Sub

Sub ReadReport()
'this sub is to start a loop to enter the oracle report window and check status of top action and loop through untill its status changes to complete
ClickOnCornerWindow
    
   
Application.SendKeys ("%v+r"), True
slow2
slow1
Application.SendKeys ("~"), True
slow2
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
slow1
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
CopyCompareCell
slow1

If Not txt1 Like "#########" Then
    'BringToFront
    'MsgBox ("Out of alignment, sub exited")
    ExitCon = True
    GoTo FINISH
End If

slow2
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
slow1

CopyCompareCell

Dim ProceedTest As Boolean
ProceedTest = True

If txt1 = "Completed" Then Application.SendKeys ("{F4}")
slow1
If txt1 = "Completed" Then ProceedTest = False
If txt1 = "Completed" Then GoTo FINISH
DoEvents


Dim i As Integer
i = 1

If ProceedTest = True Then
    For i = 1 To 21
        If ProceedTest = True Then
            Application.Wait (Now + TimeValue("00:00:04"))
            Application.SendKeys ("{Tab}")
            
            Application.SendKeys ("%r"), True
            slow1
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
            CopyCompareCell
            If i = 21 Then ExitCon = True
            Dim timetest As Boolean
            timetest = True
            DoEvents
            If ExitCon = True Then GoTo SkipToEnd
            
            If txt1 = "Completed" Then ProceedTest = False 'Else
            slow2
            
        Else
        i = 21
        End If
    
    Next i
End If

Application.SendKeys ("%fc")

GoTo FINISH

SkipToEnd:
If ExitCon = True Then Application.SendKeys ("%fc")

If timetest = True Then BringToFront
If timetest = True Then MsgBox "Request has stalled, review and then Restart Launch Cost"

ThisWorkbook.Worksheets("AIF").Cells(findCell.Row, 9) = "Stage" & " " & cycleLoop

FINISH:
End Sub
Sub setvars()
' this sub assigns value for variabled called in various subs

t = findCell.Column

Item = Cells(findCell.Row, t).Value
OrgCodeStart = Cells(findCell.Row, (t + 1)).Value
orgCode = Cells(findCell.Row, (t + 1)).Value
OrgNum = Cells(findCell.Row, t + 2).Value
CostType = Cells(findCell.Row, t + 3).Value
MorA = Cells(findCell.Row, t + 4).Value
AifType = Cells(findCell.Row, t + 5).Value

'Stop
End Sub
Sub MovetoHistory()
'this function moves all rows which contain data and are marked with a color other than no-fill to history section

Dim Cl As Range
Dim wrkRng As Range
Dim hisRng As Range
Set wrkRng = ThisWorkbook.Worksheets("AIF").Range("B5: B45 ")
For Each Cl In wrkRng

Dim findCell As Range
Dim empHist As Range
Dim rng As Range

Set hisRng = ThisWorkbook.Worksheets("AIF").Range("b50:B1048576")
Set rng = ThisWorkbook.Worksheets("AIF").Range("B5:B45")

'Application.FindFormat.Clear
'Application.FindFormat.Interior.ColorIndex = 2

Set findCell = rng.Find(what:="*", searchFormat:=False)
Application.FindFormat.Clear
Application.FindFormat.Interior.ColorIndex = -4142
Set empHist = hisRng.Find(what:="", searchFormat:=True)

If findCell Is Nothing Then BringToFront
If findCell Is Nothing Then MsgBox "Rows moved to history"
If findCell Is Nothing Then End

If findCell.Offset(0, 7) = "Completed" Then
    If Not findCell.Interior.ColorIndex = -4142 Then
        Range(findCell, findCell.Offset(0, 10)).Cut Range(empHist, empHist.Offset(0, 10))
    End If
End If

Next
End Sub


Sub Jointcmd()
' this sub loops through the list of item #s on the AIF sheet and calls subs to achieve the cost update process and marks off items as they are completed

StartOver:
If ExitCon = True Then retry1 = True
ExitCon = False
Dim Answer As String



slow1
Dim Cl As Range
Dim wrkRng As Range
Set wrkRng = ThisWorkbook.Worksheets("AIF").Range("B4:B113")
For Each Cl In wrkRng

Dim rng As Range
   
Set rng = ThisWorkbook.Worksheets("AIF").Range("B3:B45")
Application.FindFormat.Clear
Application.FindFormat.Interior.ColorIndex = 2

Set findCell = rng.Find(what:="*", searchFormat:=True)
If (findCell Is Nothing) Then
    BringToFront
    MsgBox "No Unprocessed Items"
    Exit Sub
    

End If
    
setvars

DoEvents
UpdateAcctCode
If Left(OrgCodeStart, 3) = "SLB" Then BringToFront
If Left(OrgCodeStart, 3) = "SLB" Then MsgBox "Use other script for Other ORG, this is for SLB only"
If Left(OrgCodeStart, 3) = "SLB" Then BringToFront


slow2
ClickOnCornerWindow
slow1
Dim stageNum As Integer
If stageNum = "0" Then stageNum = Empty
Dim Stagechk As String

Stagechk = ThisWorkbook.Worksheets("AIF").Cells(findCell.Row, 9)

If Stagechk Like "*[0-9]*" Then
    stageNum = Right(ThisWorkbook.Worksheets("AIF").Cells(findCell.Row, 9), 1)
    Stagechk = Left(ThisWorkbook.Worksheets("AIF").Cells(findCell.Row, 9), 5)
End If



If Stagechk = "Stage" Then cycleLoop = stageNum
slow1
If cycleLoop > 0 Then GoTo MidProgress


Application.SendKeys ("%tt"), True
slow2
changeOrg
If ExitCon = True Then GoTo SkipToEnd


slow1
Application.SendKeys ("%tt")
LaunchCost
DoEvents
If ExitCon = True Then GoTo SkipToEnd
ReadReport
DoEvents
If ExitCon = True Then GoTo SkipToEnd
slow2
ChkInhouseCost
If AifType = "Blend" Then CopyToKickForBlend
'ClickOnCornerWindow
Application.SendKeys ("%tt"), True

MidProgress:
For cycleLoop = 1 To 4
If cycleLoop < stageNum Then cycleLoop = stageNum

    incrementCycle

    UpdateAcctCode
    ThisWorkbook.Worksheets("AIF").Cells(findCell.Row, 9) = "Stage" & " " & cycleLoop
    'MsgBox orgCode
    'Stop
    slow2
    Application.SendKeys ("%tt"), True
    slow2
    changeOrg
    If ExitCon = True Then GoTo SkipToEnd
    slow1
    Application.SendKeys ("%tt"), True
    slow2
  
    CopyCost
    If ExitCon = True Then GoTo SkipToEnd
    ReadReport
    If ExitCon = True Then GoTo SkipToEnd
    slow2
    Application.SendKeys ("%tt"), True
    slow1
    LaunchCost
    slow1
    If ExitCon = True Then GoTo SkipToEnd
    ReadReport
    slow1
    If ExitCon = True Then GoTo SkipToEnd
    If AifType = "Blend" Then CopyToKickForBlend
    slow1
    
Next cycleLoop

cycleLoop = Empty
findCell.Interior.ColorIndex = 4
ThisWorkbook.Worksheets("AIF").Cells(findCell.Row, 9) = "Completed"
ThisWorkbook.Worksheets("AIF").Cells(findCell.Row, 12) = Now
retry1 = False
stageNum = 0
Next Cl

SkipToEnd:
If retry1 = False Then GoTo StartOver
Application.SendKeys ("{NUMLOCK}"), True
BringToFront
If ExitCon = True Then MsgBox "Error encountered, Stopping script"

End Sub
Sub BackTOstart()



End Sub
Sub OutSourceChk()

Dim Answer As String
Dim D1 As Range
Dim rng As Range
Dim wrkRng As Range
Set wrkRng = ThisWorkbook.Worksheets("AIF").Range("B4:B40")
Set rng = ThisWorkbook.Worksheets("AIF").Range("B4:B40")

'BringToFront
'Answer = MsgBox("Copy Transfer to XYZ Backup for Outsource", vbQuestion + vbYesNo + vbDefaultButton2, "Message Box Title")
'BringToFront

'Set findCell = rng.Find(what:="*", searchFormat:=True)
If Answer = vbYes Then
    Application.FindFormat.Clear
    Application.FindFormat.Interior.ColorIndex = 2
For Each D1 In wrkRng

    Set findCell = rng.Find(what:="Outsource", searchFormat:=True)
    
    On Error Resume Next
    If (findCell.Offset(0, -5).Interior.ColorIndex = 4) Then GoTo OutSkip
    On Error GoTo 0
    If (findCell Is Nothing) Then
        Stop
        BringToFront
        MsgBox "All Reports Pulled"
        Exit Sub
    End If
 
   
    ClickOnCornerWindow
    slow2
    Application.SendKeys "%tt"
    slow2
    Application.SendKeys "2"
    slow1
    Application.SendKeys "%o"
    slow1
    Application.SendKeys "Copy Item Costs"
    Application.SendKeys ("{Tab}")
    slow1
    Application.SendKeys "Remove"
    Application.SendKeys ("{Tab}")
    slow1
    Application.SendKeys "TRANSFER"
    Application.SendKeys ("{Tab}")
    Application.SendKeys "XYZ-BACKUP"
    slow2
    Application.SendKeys ("{Tab}")
    'Application.SendKeys ("+{Tab}")
    Application.SendKeys ("{Tab}")
    slow1
    Application.SendKeys ("%o")
    Application.SendKeys ("%o")
    slow1
    Application.SendKeys "S"
    
    Application.SendKeys ("{Tab}")
    slow2
    
    Application.SendKeys findCell.Offset(0, -5).Value
    slow2
    Application.SendKeys ("%o")
    Application.SendKeys ("%o")
    slow1
    Application.SendKeys ("%m")
    slow1
    Application.SendKeys ("%n")
    
    
    findCell.Interior.ColorIndex = 4
OutSkip:
Stop
Next

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

Application.SendKeys ("5"), True
slow1
Application.SendKeys ("+{Tab}")
CopyCompareCell
If Not txt1 = "%" Then ExitCon = True
If ExitCon = True Then GoTo Skip
Application.SendKeys ("{Tab}")

Application.SendKeys (c), True
slow1
Skip:
If ExitCon = True Then Application.SendKeys ("%c")
End Sub
Sub incrementCycle()
' this is a operation for jointcmd sub to keep track of what step it is at in the cost update process

Dim rng3 As Range
Dim rng4 As Range
Dim findOrg As Range

Dim OrgRow As Range

Dim incrementCol As Range
    
Set rng3 = ThisWorkbook.Worksheets("LocArray").Range("c5:c26")

Set findOrg = rng3.Find(what:=OrgCodeStart, MatchCase:=False)
Dim t As Integer
Dim b As Integer
t = findOrg.Column
b = findOrg.Row
orgCode = ThisWorkbook.Worksheets("LocArray").Cells(b, t + cycleLoop).Value
OrgNum = ThisWorkbook.Worksheets("LocArray").Cells(b + 1, t + cycleLoop).Value

End Sub
Sub LaunchCost()
'sub handles navigation and keystrokes for processing through the launch cost update form
Application.SendKeys ("4"), True

slow2

Application.SendKeys CostType

slow2
Application.SendKeys ("{Tab}")
slow1
Application.SendKeys (OrgNum), True
slow2
If AifType = "CriticalPart" Then
    If orgCode = "GWH" Or orgCode = "CNL" Then
    If MorA = "Assm" Then Application.SendKeys "1019"
    If MorA = "Mold" Then Application.SendKeys "1018"
    slow1
'If orgCode = "CNL" And AifType = "CriticalPart" Then
'   If MorA = "Assm" Then Application.SendKeys "1019"
'   If MorA = "Mold" Then Application.SendKeys "1018"
    Else
    Application.SendKeys (AcctCode), True
    slow1
    End If
Else
Application.SendKeys (AcctCode), True
slow1
End If

'End If


Application.SendKeys (502800), True
Application.SendKeys ("%o"), True

Application.SendKeys (orgCode)
Application.SendKeys (".")

Application.SendKeys (Item), True
slow2
Application.SendKeys ("{Tab}")

CopyCompareCell

'If Not txt1 = "Specific item" Then BringToFront
'If Not txt1 = "Specific item" Then MsgBox "Error Condition, stopping cost update"
If Not txt1 = "Specific item" Then ExitCon = True
If Not txt1 = "Specific item" Then GoTo SkipToEnd

Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")

Application.SendKeys ("r")
Application.SendKeys ("{Tab}")
slow1
Application.SendKeys (Item), True
Application.SendKeys ("%o"), True
Application.SendKeys ("%c"), True
slow1
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
CopyCompareCell
'If Not txt1 = "Yes" Then BringToFront
'If Not txt1 = "Yes" Then MsgBox "Error Condition, stopping cost update"
If Not txt1 = "Yes" Then ExitCon = True
If Not txt1 = "Yes" Then GoTo SkipToEnd

Application.SendKeys ("%o"), True
slow1

Application.SendKeys ("%m"), True
Application.SendKeys ("%n"), True

GoTo FINISH

SkipToEnd:
If ExitCon = True Then
    
    Application.SendKeys ("{ESC}")
    Application.SendKeys ("{ESC}")
    slow1
    Application.SendKeys ("{ESC}")
    Application.SendKeys ("{ESC}")
    slow1
    'Application.SendKeys ("{Esc}")
    'Application.SendKeys ("{Esc}")
    'slow1
    Application.SendKeys "%n"
    
End If
ThisWorkbook.Worksheets("AIF").Cells(findCell.Row, 9) = "Stage" & " " & cycleLoop

FINISH:
End Sub
Sub CopyToKickForBlend()
'this sub is for handling copy cost action for loading new blends
Application.SendKeys ("2") 'copy cost

DoEvents
Sleep 100
DoEvents

Application.SendKeys ("%o")

slow2
DoEvents

Application.SendKeys ("Copy Item Costs"), True
slow2
Application.SendKeys ("{Tab}")

slow2

Application.SendKeys ("R")
Sleep 100
Application.SendKeys ("{Tab}")

'Application.SendKeys (orgCode), True

Application.SendKeys ("Transfer")
Application.SendKeys ("{Tab}")
Application.SendKeys ("Kickoff")
Application.SendKeys ("{Tab}")
Application.SendKeys ("s")
Application.SendKeys ("{Tab}")

Application.SendKeys (Item), True
Application.SendKeys ("%o"), True

slow2
Application.SendKeys ("+{Tab}")
slow1
CopyCompareCell
slow1

If Not txt1 = Item Then BringToFront
If Not txt1 = Item Then MsgBox "Error Condition, stopping cost update"
If Not txt1 = Item Then ExitCon = True
If Not txt1 = Item Then GoTo SkipToEnd
slow1
Application.SendKeys ("%o")
slow1
Application.SendKeys ("%o")
Application.SendKeys ("%m")
Application.SendKeys ("%n")
slow2
Application.SendKeys ("2") 'copy cost

DoEvents
Sleep 100
DoEvents

Application.SendKeys ("%o")

slow2
DoEvents

Application.SendKeys ("Copy Item Costs"), True
slow2
Application.SendKeys ("{Tab}")

slow2

Application.SendKeys ("R")
Sleep 100
Application.SendKeys ("{Tab}")

'Application.SendKeys (orgCode), True

Application.SendKeys ("Transfer")
Application.SendKeys ("{Tab}")
Application.SendKeys ("Pending")
Application.SendKeys ("{Tab}")
Application.SendKeys ("s")
Application.SendKeys ("{Tab}")

Application.SendKeys (Item), True
Application.SendKeys ("%o"), True

slow2
Application.SendKeys ("+{Tab}")
slow1
CopyCompareCell
slow1

If Not txt1 = Item Then BringToFront
If Not txt1 = Item Then MsgBox "Error Condition, stopping cost update"
If Not txt1 = Item Then ExitCon = True
If Not txt1 = Item Then GoTo SkipToEnd
slow1
Application.SendKeys ("%o")
slow1
Application.SendKeys ("%o")
Application.SendKeys ("%m")
Application.SendKeys ("%n")
GoTo FINISH

SkipToEnd:
ThisWorkbook.Worksheets("AIF").Cells(findCell.Row, 9) = "Stage" & " " & cycleLoop


FINISH:
End Sub
Sub CopyCost()
' this sub is for handling navigation and keystrokes for copy cost form
Application.SendKeys ("2") 'copy cost

DoEvents
Sleep 100
DoEvents

Application.SendKeys ("%o")

slow2
DoEvents

Application.SendKeys ("Copy Item Costs A"), True
slow2
Application.SendKeys ("{Tab}")

slow2

Application.SendKeys ("R")
slow2
Application.SendKeys ("{Tab}")
Application.SendKeys ("+{Tab}")
CopyCompareCell
'If Not txt1 = "Remove and replace all cost information" Then BringToFront
'If Not txt1 = "Remove and replace all cost information" Then MsgBox "Error Condition, stopping cost update"
If Not txt1 = "Remove and replace all cost information" Then ExitCon = True
If Not txt1 = "Remove and replace all cost information" Then GoTo SkipToEnd
slow2
Application.SendKeys ("{Tab}")

Application.SendKeys (OrgCodeStart), True

Application.SendKeys (CostType), True
Application.SendKeys ("{Tab}")
Application.SendKeys (CostType), True
Application.SendKeys ("{Tab}")
Application.SendKeys ("s")
Application.SendKeys ("{Tab}")

Application.SendKeys (Item), True
Application.SendKeys ("%o"), True

slow2
Application.SendKeys ("cr")
slow1
Application.SendKeys ("{Tab}"), True
slow2

CopyCompareCell
'If Not txt1 = "Corporate" Then BringToFront
'If Not txt1 = "Corporate" Then MsgBox "Error Condition, stopping cost update"
If Not txt1 = "Corporate" Then ExitCon = True
If Not txt1 = "Corporate" Then GoTo SkipToEnd

Application.SendKeys ("%o")
Application.SendKeys ("%m")
Application.SendKeys ("%n")
GoTo FINISH

SkipToEnd:
If ExitCon = True Then
    Application.SendKeys ("{ESC}")
    Application.SendKeys ("{ESC}")
    slow1
    Application.SendKeys ("{ESC}")
    Application.SendKeys ("{ESC}")
    slow1
    Application.SendKeys "%n"
    slow1
    Application.SendKeys "%c"
End If
ThisWorkbook.Worksheets("AIF").Cells(findCell.Row, 9) = "Stage" & " " & cycleLoop

FINISH:
End Sub
Sub UpdateAcctCode()
' this is a sub for other subs to call when acct codes need updated between steps
Dim rng1 As Range
Dim rng2 As Range
Dim findCellRow As Range
Dim OrgCol As Range
Set rng1 = ThisWorkbook.Worksheets("LocArray").Range("i4:i14")
Set rng2 = ThisWorkbook.Worksheets("LocArray").Range("j3:n3")

Set findCellRow = rng1.Find(what:=AifType, MatchCase:=False)
If findCellRow Is Nothing Then
    Set findCellRow = rng1.Find(what:=MorA, MatchCase:=False)
End If

Set OrgCol = rng2.Find(what:=orgCode, MatchCase:=False)

DoEvents

If OrgCol = "" Then MsgBox "Please Select a Item #"
If OrgCol Is Nothing Then ExitCon = True

If OrgCol Is Nothing Then Exit Sub

AcctCode = Worksheets("LocArray").Cells(findCellRow.Row, OrgCol.Column).Value
'Stop
End Sub
Sub RepeatReportReader()
'this sub cycles through the item list on the AIF sheet and calls the pull report subs for each item and registers them as yellow when finished
ExitCon = False

Dim Outsourcecheck As Boolean
Dim C1 As Range
Dim wrkRng As Range
Set wrkRng = ThisWorkbook.Worksheets("AIF").Range("B4:B40")
For Each C1 In wrkRng

    Dim rng As Range

    Set rng = ThisWorkbook.Worksheets("AIF").Range("B4:B40")
    Application.FindFormat.Clear
    Application.FindFormat.Interior.ColorIndex = 2
    Set findCell = rng.Find(what:="*", searchFormat:=True)
    On Error Resume Next
    If findCell.Offset(0, 5) = "Outsource" Then Outsourcecheck = True
    On Error GoTo 0
    If (findCell Is Nothing) Then
        GoTo check
        Exit Sub
    End If
    
    setvars
    If AifType = "Outsource" And MorA = "Assm" Then BringToFront
    If AifType = "Outsource" And MorA = "Assm" Then MsgBox "ADD 7.5% for FG, also double check cost and BOM is copied to proper ORG"
    slow1
    DoEvents
    UpdateAcctCode
    
    slow2
    ClickOnCornerWindow
    
    
    slow2
    Application.SendKeys ("%w1"), True
    changeOrg
    slow1
    pullCostReports
    If ExitCon = True Then GoTo SkipToEnd
    
    slow2
    ClickOnCornerWindow
    findCell.Interior.ColorIndex = 6
    ThisWorkbook.Worksheets("AIF").Cells(findCell.Row, 9) = "Report Called"
    ThisWorkbook.Worksheets("AIF").Cells(findCell.Row, 15) = WorksheetFunction.Concat(Item, " Custom Cost Sheet ", CostType, " ", WorksheetFunction.Text(Date, "mm-dd-yy"))
    ThisWorkbook.Worksheets("AIF").Cells(findCell.Row, 17) = WorksheetFunction.Concat(Item, " Supply Chain Cost Rollup ", CostType, " ", WorksheetFunction.Text(Date, "mm-dd-yy"))
    slow2
    ChkInhouseCost
'SkipToNext:

Next C1

check:
'If Outsourcecheck = True Then OutSourceChk
BringToFront
MsgBox "All Reports Pulled"


SkipToEnd:
Application.SendKeys ("{NUMLOCK}"), True
End Sub

Sub BringToFront()
'this sub bring the AIF sheet into focus
    Dim setFocus As Long
    
    ThisWorkbook.Worksheets("AIF").Activate
    setFocus = SetForegroundWindow(Application.hwnd)
End Sub

Sub endProcess()
'this sub is for the cost update processes to stop but attempt to log it process before exit.
If cycleLoop = Empty Then GoTo Skip
ThisWorkbook.Worksheets("AIF").Cells(findCell.Row, 9) = "Stage" & " " & cycleLoop

DoEvents
Skip:
End
End Sub
Sub AddPeiceWeight()
'sub for adding peice weight to master item at item transfer
ClickOnCornerWindow
'If Cells(findCell.Row, 9) > "0" Then
If 0 = 0 Then
    Application.SendKeys ("%fw")
    slow2
    Application.SendKeys ("ni")
    slow2
    Application.SendKeys ("%tt")
    slow1
    Application.SendKeys ("1")
    slow1
    Application.SendKeys ("cc")
    slow2
    Application.SendKeys Item
    slow2
    Application.SendKeys ("{Tab}")
    slow1
    Application.SendKeys ("+{Tab}")
    CopyCompareCell
    If Not txt1 = "Item" Then
    If Not txt1 = "Item" Then MsgBox "out of alignment, please restart"
    If Not txt1 = "Item" Then End
    Application.SendKeys ("{Tab}")
    Application.SendKeys ("{Tab}")
    Application.SendKeys ("{Tab}")
    Application.SendKeys ("{Tab}")
    slow2
    Application.SendKeys ("{Tab}")
    Application.SendKeys ("{Tab}")
    Application.SendKeys ("{Tab}")
    Application.SendKeys ("{Tab}")
    slow2
    Application.SendKeys ("{Tab}")
    Application.SendKeys ("{Tab}")
    Application.SendKeys ("{Tab}")
    Application.SendKeys ("{Tab}")
    slow2
    Application.SendKeys ("{Tab}")
    Application.SendKeys ("{Tab}")
    Application.SendKeys ("{Tab}")
    Application.SendKeys ("{Tab}")
    slow2
    Application.SendKeys ("{Tab}")
    slow1
    Application.SendKeys "99"
    'Application.SendKeys Cells(findCell.Row, 8).Value
    slow1
    Application.SendKeys ("{Tab}")
    Application.SendKeys ("{Tab}")
    Application.SendKeys ("{Tab}")
    Application.SendKeys ("{Tab}")
    slow2
    CopyCompareCell
    If Not txt1 = "BOTH" Then BringToFront
    If Not txt1 = "BOTH" Then MsgBox "out of alignment, please restart"
    If Not txt1 = "BOTH" Then End
    slow1
    Application.SendKeys ("%fv")
    slow2
    Application.SendKeys ("{f4}")
End If
End If
Application.SendKeys ("%fw")
slow2
Application.SendKeys ("cc")
slow2

End Sub

Sub MassCopy()
'sub for mass copy sheet, needs updating and better explination on use, currenly funcitioning well however
'generaly used when we need to update pending or kickoff cost for mass list of items

ClickOnCornerWindow
ExitCon = False

'set window range to right halfscreen
Dim Mzx As Range
Dim wRng As Range
Set wRng = ThisWorkbook.Worksheets("MassCopy").Range("B4:B113")
For Each Mzx In wRng

Dim rng As Range
   
Set rng = ThisWorkbook.Worksheets("MassCopy").Range("B4:B45")
Application.FindFormat.Clear
Application.FindFormat.Interior.ColorIndex = 2

Set findCell = rng.Find(what:="*", searchFormat:=True)
If (findCell Is Nothing) Then
    BringToFront
    MsgBox "No Unprocessed Items"
    Exit Sub


End If

Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
slow2
Application.SendKeys ("{Tab}")
CopyCompareCell
If Not txt1 = "Specific item" Then GoTo SkipToEnd
slow2
Application.SendKeys ("{Tab}")
Application.SendKeys Cells(findCell.Row, 2).Value
Application.SendKeys ("%oo")
slow1
Application.SendKeys ("%m")
Application.SendKeys ("%y")
slow2


findCell.Interior.ColorIndex = 4
ThisWorkbook.Worksheets("MassCopy").Cells(findCell.Row, 9) = "Completed"
txt1 = "0"
Next Mzx

SkipToEnd:
Dim setFocus As Long
    
ThisWorkbook.Worksheets("MassCopy").Activate

End Sub

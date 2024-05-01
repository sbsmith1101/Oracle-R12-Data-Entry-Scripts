Attribute VB_Name = "ReplaceCostCode"

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
Public ItemNum As String
Public t As String

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
Sub BringToFront()
'brings kickoff boms sheet into focus
    Dim setFocus As Long
    
    ThisWorkbook.Worksheets("Replace cost elements").Activate
    setFocus = SetForegroundWindow(Application.hWnd)
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
Sub MainSeq()

slow1
ClickOnCornerWindow
slow1
'Application.SendKeys ("%tc")
'Application.SendKeys ("%tu")
'Application.SendKeys ("i")
'Application.SendKeys ("~")

slow1
Dim Cl As Range
Dim wrkRng As Range
Set wrkRng = ThisWorkbook.Worksheets("Replace cost elements").Range("c1:c1000")
For Each Cl In wrkRng

Dim rng As Range
   
Set rng = ThisWorkbook.Worksheets("Replace cost elements").Range("c1:c1000")
Application.FindFormat.Clear
Application.FindFormat.Interior.ColorIndex = 2

Set findCell = rng.Find(what:="*", searchFormat:=True)
If (findCell Is Nothing) Then
    BringToFront
    MsgBox "No Unprocessed Items"
    Exit Sub
    

End If

ClickOnCornerWindow
If Not ItemNum = Left(findCell, InStr(findCell, ",") - 1) Then NewITEM

ItemNum = Left(findCell, InStr(findCell, ",") - 1)



Application.SendKeys ("%fv")
'Application.SendKeys ("%vf")


Next Cl


End Sub
Sub NewITEM()
ItemNum = Left(findCell, InStr(findCell, ",") - 1)
Application.SendKeys ("%vf")
Application.SendKeys ItemNum
Application.SendKeys ("{Tab}")
Application.SendKeys "2023-SLB"
'Application.SendKeys "2022-ACM"'for testing
slow1
Application.SendKeys ("%i")
Application.SendKeys ("%c")
slow1
slow1
CopyCompareCell

If Not InStr(txt1, "Material") = 1 Then BringToFront
If Not InStr(txt1, "Material") = 1 Then MsgBox "Skipping Detected pleas re-start item"
If Not InStr(txt1, "Material") = 1 Then Stop
For a = 0 To 8
    If Not InStr(txt1, "Material") = 1 Then GoTo skip1
    Application.SendKeys ("%ed")
    Application.SendKeys ("%o")
    'Application.SendKeys ("{Down}")
    slow1
    CopyCompareCell
skip1:
Next

slow1

For x = 0 To 8
    If Not Left(findCell.Offset(x, 0), InStr(findCell, ",") - 1) = ItemNum Then GoTo Skip
    
    findCell.Offset(0 + x, 1) = Mid(findCell.Offset(x, 0), InStr(findCell, ",") + 1)
    Application.SendKeys Left(findCell.Offset(0 + x, 1), InStr(findCell.Offset(0 + x, 1), ",") - 1)
    Application.SendKeys ("{Tab}")
    slow1
    findCell.Offset(0 + x, 1) = Mid(findCell.Offset(0 + x, 1), InStr(findCell.Offset(0 + x, 1), ",") + 1)
    Application.SendKeys Left(findCell.Offset(0 + x, 1), InStr(findCell.Offset(0 + x, 1), ",") - 1)
    'Application.SendKeys "Admin1"'for testing
    Application.SendKeys ("{Tab}")
    slow1
    findCell.Offset(0 + x, 1) = Mid(findCell.Offset(0 + x, 1), InStr(findCell.Offset(0 + x, 1), ",") + 1)
    Application.SendKeys ("{Tab}")
    slow1
    Application.SendKeys ("{Tab}")
    Application.SendKeys findCell.Offset(0 + x, 1)
    Application.SendKeys ("{Down}")
    findCell.Offset(0 + x, 0).Interior.ColorIndex = 4
    
    
Skip:

Next

slow1



End Sub
Sub ENDIT()

End

End Sub


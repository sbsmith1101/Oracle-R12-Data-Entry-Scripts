Attribute VB_Name = "loadSHLDres"
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


Public txt1 As String
Public opRng As Range
Public findCell As Range
Public orgCode As String
Public ItemNum As String
Public toolNum As String
Public ASSYITEM As String
Public t As String
Public PPH As Integer

Sub cancelPROP()
End
End Sub


Sub SHLDLoad()


'this sub runs a loop through the "ADD SCHD RES" sheets item list, enterting values in master item, routing, Bom, and cost elements


Dim Cl As Range
Dim wrkRng As Range
Set wrkRng = ThisWorkbook.Worksheets("ADD SCHD RES").Range("B5:B113")
For Each Cl In wrkRng


Application.FindFormat.Clear
Application.FindFormat.Interior.ColorIndex = 2

Dim rng As Range
Set rng = ThisWorkbook.Worksheets("ADD SCHD RES").Range("B5:B113")
Set findCell = rng.Find(what:="*", searchFormat:=True)
If (findCell Is Nothing) Then
    BringToFront
    MsgBox "No Unprocessed Items"
    Exit Sub
    

End If

'' trying to figure out how to compare date values so I can spot check for BOM window alignment


ClickOnCornerWindow
setvars


'open Routings
slow2
Application.SendKeys ("%fw"), True

slow1
Application.SendKeys ("nb"), True
slow1


changeBM

slow1
Application.SendKeys ("1"), True
slow1
Application.SendKeys ("%vf"), True
'enter item and check 00 to ensure we don't encouter item alread exsist error
slow1
Application.SendKeys ItemNum
Application.SendKeys ("%i"), True
Application.SendKeys ("%o"), True
'load routing details
slow2
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
Application.SendKeys ("+{pgdn}")
Application.SendKeys ("{down}")
slow2
CopyCompareCell

If Not txt1 = "20" Then BringToFront
If Not txt1 = "20" Then MsgBox "out of alignment"
If Not txt1 = "20" Then End
DoEvents
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
slow2
Application.SendKeys ("SH")
Application.SendKeys ("{Tab}")
'load shld res
slow1
Application.SendKeys "%r"
Application.SendKeys "10"
slow1
Application.SendKeys ("{Tab}")
Application.SendKeys toolNum
Application.SendKeys ("{Tab}")
CopyCompareCell

If Not txt1 = "HR" Then BringToFront
If Not txt1 = "HR" Then MsgBox "out of alignment"
If Not txt1 = "HR" Then End
DoEvents
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
slow1
Application.SendKeys PPH
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
slow1
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
slow1
Application.SendKeys "y"
Application.SendKeys ("{Tab}")
slow1
Application.SendKeys ("{down}")
slow1
Application.SendKeys "20"
slow1
Application.SendKeys ("{Tab}")
Application.SendKeys ASSYITEM
Application.SendKeys ("{Tab}")
CopyCompareCell

If Not txt1 = "HR" Then BringToFront
If Not txt1 = "HR" Then MsgBox "out of alignment"
If Not txt1 = "HR" Then End
DoEvents
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
slow1
Application.SendKeys PPH
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
slow1
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
slow1
Application.SendKeys "y"
Application.SendKeys ("{Tab}")
slow1
'ClickmainRoutBox

Application.SendKeys "%fs"
slow2
Application.SendKeys "%fc"

findCell.Interior.ColorIndex = 5
Next
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
Sub BringToFront()
'brings kickoff boms sheet into focus
    Dim setFocus As Long
    
    ThisWorkbook.Worksheets("ADD SCHD RES").Activate
    setFocus = SetForegroundWindow(Application.hwnd)
End Sub
Sub changeBM()
'this is for changing orgs in Bom and routing section
'orgCode = Cells(iRow, 30).Value
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
Sub setvars()
'this sub sets assigns the values of item variables which are used by other subs
t = findCell.Column

ItemNum = Cells(findCell.Row, 2).Value
toolNum = Cells(findCell.Row, 3).Value
ASSYITEM = Cells(findCell.Row, 4).Value
orgCode = Cells(findCell.Row, 6).Value
PPH = Round(Cells(findCell.Row, 5).Value, 0)

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

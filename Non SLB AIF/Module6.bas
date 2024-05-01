Attribute VB_Name = "Module6"
Public Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As LongPtr, ByVal y As LongPtr) As LongPtr
Public Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwmilliseconds As LongPtr)

Public Declare PtrSafe Sub mouse_event Lib "user32" (ByVal dwFlags As LongPtr, ByVal dx As Long, ByVal dy As LongPtr, ByVal cButtons As LongPtr, ByVal swextrainfo As LongPtr) '
Public Const mouseeventf_Leftdown = &H2
Public Const mouseeventf_Leftup = &H4
Public Const mouseeventf_Rightdown As Long = &H8
Public Const mouseeventf_rightup As Long = &H10
Sub searchFindFormat()
    Dim rng As Range
    Dim findCell As Range
    
Set rng = ThisWorkbook.Worksheets("Test").Range("B13:B35")
Application.FindFormat.Clear
Application.FindFormat.Interior.ColorIndex = 2

Set findCell = rng.Find(what:="*", searchFormat:=True)

If (findCell Is Nothing) Then
    MsgBox keyword & "Not Found"
    'Range("B13").Interior.ColorIndex = 4

Else
    MsgBox findCell.Row
    
End If

End Sub
Sub colormepurple()

Cells(18, 8).Interior.ColorIndex = 12
End Sub

Sub showcell()



Dim Item As String
Item = Cells(13, 2).Value
  
Dim orgCode As String
orgCode = Range("N34").Value

Dim OrgNum As String
OrgNum = Range("N34").Value

Dim CostType As String
CostType = Range("O34").Value

Dim AcctCode As String
AcctCode = Range("O34").Value

MsgBox Item


End Sub

Sub CodeFinder()

Dim rng1 As Range
Dim rng2 As Range
Dim findCellCol As Range
Dim MorA As String
MorA = "assm"
Dim AifType As String
AifType = "transfer"
Dim Org As String
Org = "MEX"
Dim OrgRow As Range
Dim CellCode As String
    
    
Set rng1 = ThisWorkbook.Worksheets("LocArray").Range("i4:i12")
Set rng2 = ThisWorkbook.Worksheets("LocArray").Range("j3:n3")

Set findCellCol = rng1.Find(what:=AifType, MatchCase:=False)
If findCellCol Is Nothing Then
Set findCellCol = rng1.Find(what:=MorA, MatchCase:=True)
Else
End If
Set OrgRow = rng2.Find(what:=Org, MatchCase:=False)

'Set CellCode = Cells(OrgRow, findCellCol).Value

MsgBox Cells(findCellCol.Row, OrgRow.Column).Value
Stop



End Sub
Sub changeColor()

ThisWorkbook.Worksheets("Test").Cells(14, 9).Value = "Completed"

End Sub

Sub copylist()

'slow2
ClickOnCorner1Window
Application.Wait (Now + TimeValue("00:00:01"))

Dim Mer As Range
'Mer = ThisWorkbook.Sheets(Sheet1).Range("h16:h281")
Dim rowI As Integer
rowI = 9
For rowI = 1327 To 1346



'MsgBox (copyCell)
slowMill
Application.SendKeys (", ")
Application.SendKeys ("<")
slowMill
Application.SendKeys ("<")
Application.SendKeys ThisWorkbook.Worksheets("KickOFF").Cells(rowI, 1).Value
slowMill
Application.SendKeys ("<")
Application.SendKeys ("~")
Next rowI
    
End Sub
Sub slowMill()
DoEvents
Application.Wait (Now + TimeValue("00:00:01") / 2)
   
End Sub

Public Sub ClickOnCorner1Window()

Dim oLeft As Long
oLeft = 100  'last desktop size
oLeft = 80   'revised desktop

Dim OTop As Long
OTop = 100

SetCursorPos oLeft, OTop

mouse_event mouseeventf_Leftdown, 0, 0, 0, 0
mouse_event mouseeventf_Leftup, 0, 0, 0, 0


End Sub
Sub BringToFront()
    Dim setFocus As Long
    
    ThisWorkbook.Worksheets("database").Activate
    setFocus = SetForegroundWindow(Application.hwnd)
End Sub

Sub f5()
ClickOnCornerWindow
Application.SendKeys ("{F5}")
End Sub
Sub findlastWordinColumn()

Dim lRow As Long

lRow = Cells(Rows.Count, 2).End(xlUp).Row

MsgBox lRow

End Sub


Attribute VB_Name = "AlphaCode"
Public Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As LongPtr, ByVal y As LongPtr) As LongPtr
Public Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwmilliseconds As LongPtr)

Public Declare PtrSafe Sub mouse_event Lib "user32" (ByVal dwFlags As LongPtr, ByVal dx As Long, ByVal dy As LongPtr, ByVal cButtons As LongPtr, ByVal swextrainfo As LongPtr) '
Public Const mouseeventf_Leftdown = &H2
Public Const mouseeventf_Leftup = &H4
Public Const mouseeventf_Rightdown As Long = &H8
Public Const mouseeventf_rightup As Long = &H10


Option Explicit

Public ItemNum As String
Public Cost As String
Public AlphaNum As String
Public findCell As Range
Public txt1 As String
Public t As String
Public Org As String



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


Sub CopyCompareCell()
'this sub activates a crtl C action and then pulls the info in the windows clipboard and stores it in a variable so it can be evaluated at various intervals
Dim ClipObj As New DataObject

ClipObj.SetText Text:=Empty
ClipObj.PutInClipboard

Application.SendKeys ("^c")
slow1



'Need too add clear clipboard Function here

On Error Resume Next
ClipObj.GetFromClipboard
On Error GoTo 0
On Error Resume Next
    txt1 = ClipObj.GetText(1)
On Error GoTo 0

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
Sub BringToFront1()
'brings kickoff boms sheet into focus
    Dim setFocus As Long
    
    ThisWorkbook.Worksheets("ALPHA PARTS WIP").Activate
    setFocus = SetForegroundWindow(Application.hwnd)
End Sub

Sub CopyOrgMasterInfo()

Dim Cl As Range
Dim wrkRng As Range
Set wrkRng = ThisWorkbook.Worksheets("ALPHA PARTS WIP").Range("B2:B113")


Application.FindFormat.Clear
Application.FindFormat.Interior.ColorIndex = 2

Dim rng As Range
Set rng = ThisWorkbook.Worksheets("ALPHA PARTS WIP").Range("B2:B113")
Set findCell = rng.Find(what:="*", searchFormat:=True)
If (findCell Is Nothing) Then
    BringToFront1
    MsgBox "No Unprocessed Items"
    Exit Sub
    

End If
setvars
slow2


ClickOnCornerWindow


If Cells(findCell.Row, 11) = "" Then MsgBox "Org Code is needed, Org field is blank"
If Cells(findCell.Row, 11) = "" Then HALT

Application.SendKeys "%fw"
slow2
Application.SendKeys "ni"
slow2
Application.SendKeys "%tt"
Application.SendKeys "1"
slow1
Application.SendKeys "CC"
slow1
Application.SendKeys "%vf"
slow2
Application.SendKeys ItemNum
Application.SendKeys "%i"
Application.SendKeys ("{Tab}")
slow1
CopyCompareCell

slow1
Cells(findCell.Row, 4) = txt1
If Cells(findCell.Row, 4) = "" Then MsgBox "Description didn't copy"
If Cells(findCell.Row, 4) = "" Then BringToFront1
If Cells(findCell.Row, 4) = "" Then End
slow2

Application.SendKeys "%tf"
Application.SendKeys "user"
slow1
CopyCompareCell
Cells(findCell.Row, 5) = txt1

Application.SendKeys "%ta"

slow1
Dim i As Integer
i = 1
CopyCompareCell
For i = 1 To 5
    
    CopyCompareCell
    If txt1 = "Car Line" Then
        Application.SendKeys ("{Tab}")
        CopyCompareCell
        Cells(findCell.Row, 6) = txt1
    End If
    If txt1 = "Design Ownership" Then
        Application.SendKeys ("{Tab}")
        CopyCompareCell
        Cells(findCell.Row, 7) = txt1
    End If
    If txt1 = "Inventory" Then
        Application.SendKeys ("{Tab}")
        CopyCompareCell
        Cells(findCell.Row, 8) = txt1
    End If
    If txt1 = "Standard Cost Owner" Then
        Application.SendKeys ("{Tab}")
        CopyCompareCell
        Cells(findCell.Row, 9) = txt1
        Org = Right(txt1, 3)
    End If
    If txt1 = "" Then i = 6
    txt1 = ""
    Application.SendKeys ("{Tab}")
    slow1
    CopyCompareCell
Next
Application.SendKeys "%tg"
CopyCompareCell
Cells(findCell.Row, 10) = Left(txt1, 2)

Application.SendKeys "%fc"
slow2
NewMaster
CopyRoutingBOM
End Sub
Sub CopyCElements()


Set opRng = Worksheets("ALPHA PARTS WIP").Range("B2:B35")
Application.FindFormat.Clear
Application.FindFormat.Interior.ColorIndex = 2

Dim rng As Range
Set rng = ThisWorkbook.Worksheets("ALPHA PARTS WIP").Range("B2:B113")
Set findCell = rng.Find(what:="*", searchFormat:=True)
If (findCell Is Nothing) Then
    BringToFront1
    MsgBox "No Unprocessed Items"
    Exit Sub
    

End If
setvars
ClickOnCornerWindow

Application.SendKeys "%tt"
Application.SendKeys "%fw"
slow1
Application.SendKeys "cc"
slow1
Application.SendKeys "%tt"
slow1
changeOrg

Application.SendKeys "1"

Application.SendKeys ItemNum
Application.SendKeys ("{Tab}")
Application.SendKeys "Frozen"
Application.SendKeys "%i"
slow2
slow1
Application.SendKeys "%c"
Application.SendKeys "%o"

Dim MyList As Object
Set MyList = CreateObject("Scripting.Dictionary")

MyList.Add "Material Overhead", 1
MyList.Add "Material", 2
MyList.Add "Resource", 3
Dim p As Integer
Dim Ccode As String

p = 1
For p = 1 To 8
CostR:
CopyCompareCell
If Not MyList.Exists(txt1) = True Then GoTo Scost
Ccode = Ccode & txt1 & ","
Application.SendKeys ("{Tab}")
CopyCompareCell
If Left(txt1, 5) = "Rylty" Then
    Ccode = Ccode & txt1 & "," & Round(Cost * 0.03, 5) & ","
    
    Application.SendKeys ("{Tab}")
    Application.SendKeys ("{Tab}")
    Application.SendKeys ("{Tab}")
    Application.SendKeys ("{Tab}")
    Application.SendKeys ("{Tab}")
    p = p + 1
    GoTo CostR
End If
Ccode = Ccode & txt1 & ","
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
CopyCompareCell
Ccode = Ccode & txt1 & ","
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")

Next

Scost:

Cells(findCell.Row, 16) = Ccode
Application.SendKeys "%o"
Application.SendKeys "%fc"
enterNewCost

End Sub
Sub enterNewCost()
slow1
Application.SendKeys "1"
Application.SendKeys AlphaNum
slow1
Application.SendKeys ("{down}")
Application.SendKeys AlphaNum
Application.SendKeys "%i"
slow2
Application.SendKeys ("{down}")
Application.SendKeys AlphaNum
Application.SendKeys ("{Tab}")
Application.SendKeys "pending"
slow2
Application.SendKeys "%c"


Dim p As Integer

p = 1
For p = 1 To 8
If Cells(findCell.Row, 16) = "" Then GoTo EndSpot

'MsgBox InStr(Cells(findCell.Row, 16), ",")

'MsgBox Left(Cells(findCell.Row, 16), (InStr(Cells(findCell.Row, 16), ",") - 1))
Application.SendKeys Left(Cells(findCell.Row, 16), (InStr(Cells(findCell.Row, 16), ",") - 1))
Cells(findCell.Row, 16) = Mid(Cells(findCell.Row, 16), InStr(Cells(findCell.Row, 16), ",") + 1)
Application.SendKeys ("{Tab}")
Application.SendKeys Left(Cells(findCell.Row, 16), (InStr(Cells(findCell.Row, 16), ",") - 1))
Cells(findCell.Row, 16) = Mid(Cells(findCell.Row, 16), InStr(Cells(findCell.Row, 16), ",") + 1)
slow2
Application.SendKeys ("{Tab}")
slow2
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")
Application.SendKeys Left(Cells(findCell.Row, 16), (InStr(Cells(findCell.Row, 16), ",") - 1))
Cells(findCell.Row, 16) = Mid(Cells(findCell.Row, 16), InStr(Cells(findCell.Row, 16), ",") + 1)
slow2
Application.SendKeys ("{down}")

Next

EndSpot:
slow2
Application.SendKeys "%fs"
Application.SendKeys "%fc"

End Sub

Sub NewMaster()

slow1
Application.SendKeys "%tt"
Application.SendKeys "1"
slow1
Application.SendKeys AlphaNum
Application.SendKeys ("{Tab}")
slow1
Application.SendKeys "%tc"
Application.SendKeys ("{Tab}")
slow1
Application.SendKeys ItemNum
Application.SendKeys "%a"
Application.SendKeys "%d"
slow1
Application.SendKeys "%fs"
slow1
slow1
slow1
slow1
slow1
slow1
Application.SendKeys "%to"
Application.SendKeys "%n"
slow1
Application.SendKeys "%fs"
slow1
Application.SendKeys "%ta"
Application.SendKeys ("{Tab}")
Application.SendKeys Cells(findCell.Row, 8).Value
Application.SendKeys ("{Tab}")
Application.SendKeys "Car Line"
Application.SendKeys ("{Tab}")
Application.SendKeys Cells(findCell.Row, 6).Value
Application.SendKeys ("{Tab}")
slow1
Application.SendKeys "Design Ownership"
Application.SendKeys ("{Tab}")
Application.SendKeys Cells(findCell.Row, 7).Value
Application.SendKeys ("{Tab}")
If Cells(findCell.Row, 9).Value = "" Then GoTo Skip
Application.SendKeys "Standard Cost owner"
Application.SendKeys ("{Tab}")
Application.SendKeys Cells(findCell.Row, 9).Value
Skip:
slow1
Application.SendKeys "%tg"
Application.SendKeys Cells(findCell.Row, 10).Value
Application.SendKeys "%fs"
slow2
slow2
slow2
Application.SendKeys "%o"
Application.SendKeys "%o"
slow2
Application.SendKeys "%fc"
'Org


End Sub
Sub CopyRoutingBOM()
Dim Cl As Range
Dim wrkRng As Range
Set wrkRng = ThisWorkbook.Worksheets("ALPHA PARTS WIP").Range("B2:B113")
Dim txt3 As String
txt3 = DateValue(WorksheetFunction.Text(Date, "mm/dd/yyyy"))
Application.FindFormat.Clear
Application.FindFormat.Interior.ColorIndex = 2

Dim rng As Range
Set rng = ThisWorkbook.Worksheets("ALPHA PARTS WIP").Range("B2:B113")
Set findCell = rng.Find(what:="*", searchFormat:=True)
If (findCell Is Nothing) Then
    BringToFront1
    MsgBox "No Unprocessed Items"
    Exit Sub
    

End If
setvars
ClickOnCornerWindow
slow1
ClickOnCornerWindow
Application.SendKeys "%tt"
Application.SendKeys "%fw"
slow1
Application.SendKeys "nb"
slow1
Application.SendKeys "%tt"
slow1
changeBM

Application.SendKeys "1"
slow1
Application.SendKeys AlphaNum

Application.SendKeys ("{Tab}")
Application.SendKeys "%tc"
Application.SendKeys ("{Tab}")
Application.SendKeys ItemNum
slow1
Application.SendKeys "%c"
slow2
slow2
Application.SendKeys "%o"
'Application.SendKeys "%fs"
slow1

Application.SendKeys "%tb"
Application.SendKeys AlphaNum
Application.SendKeys ("{Tab}")
Application.SendKeys "%tc"
Application.SendKeys ("{Tab}")
Application.SendKeys ItemNum
slow1
Application.SendKeys "%c"
slow2
slow2
Application.SendKeys "%o"


Application.SendKeys "%fs"
slow2
Application.SendKeys "%fc"
slow2
Application.SendKeys "%fs"
Application.SendKeys "%fc"
slow1
Application.SendKeys "%tu"
slow2

Application.SendKeys "bb"
slow1
Application.SendKeys "%vf"
slow1
Application.SendKeys AlphaNum
Application.SendKeys ("~")
Application.SendKeys "+{PGDN}"

'If Cells(findCell.Row, 15) = "TRUE" Then GoTo BOMCOPY

slow1
CopyCompareCell
Dim itemseq As Integer
Dim itemI As Integer


itemseq = txt1
slow2
Application.SendKeys "{PGUP}"
Application.SendKeys "{PGUP}"
Application.SendKeys "{PGUP}"
Application.SendKeys "{PGUP}"
slow1
Application.SendKeys ("{Tab}")
Application.SendKeys ("{Tab}")

CopyCompareCell
Dim N As Integer
Dim lst As String
N = 1
For N = 1 To 20
    If txt1 = lst Then N = 21
    If txt1 = "" Then N = 21
    If txt1 = "" Then GoTo Nxt
    If Left(txt1, 1) = "N" Then Application.SendKeys "%ed"
    If Left(txt1, 1) = "R" Then Application.SendKeys "%ed"
    If Left(txt1, 1) = "B" Then Application.SendKeys "%ed"
    If Left(txt1, 1) = "X" Then Application.SendKeys "%ed"
    If Left(txt1, 1) Like "*[1-9]*" Then Application.SendKeys "{Down}"
    lst = txt1
    
    CopyCompareCell
Nxt:
Next

'Application.SendKeys "%fs"
'add new packaging loop
Dim packArray As Variant
Dim usePArray As Variant

packArray = Split(Cells(findCell.Row, 13).Value, Chr(10))
usePArray = Split(Cells(findCell.Row, 14).Value, Chr(10))
Dim L3 As Integer
Dim I3 As Integer
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
    'ClearClipboard
    txt1 = ""
    Dim i As Integer
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
    
    slow1
    Application.SendKeys ("+{Tab}")
    slow1
    Application.SendKeys ("+{Tab}")
    slow1
    
    Application.SendKeys usePArray(I3 - 1)
    Application.SendKeys ("{Tab}")
    Application.SendKeys ("+{Tab}")
    slow1
    Stop
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
        If Cells(findCell.Row, 30) = "SLB" Then
        Application.SendKeys ("{Tab}")
        slow1
        Application.SendKeys "SLB PROD"
        End If
    End If
    slow1
Application.SendKeys ("{down}")
Stop
txt1 = 0
Next
Application.SendKeys ("{up}")
CopyCompareCell
If Not txt1 = itemseq Then EndBOM
BOMCOPY:
Stop
Application.SendKeys "%fs"
slow1
Application.SendKeys "%fc"
slow2
Application.SendKeys "%fs"
slow1
Application.SendKeys "%fc"
Stop

End Sub
Sub HALT()
End
End Sub
Sub changeBM()
'this is for changing orgs in Bom and routing section
Org = Cells(findCell.Row, 11).Value
slow1

Application.SendKeys ("3"), True
slow1

Dim c As String
If Org = "CNL" Then c = "cn"
If Org = "GWH" Then c = "g"
If Org = "LVG" Then c = "l"
If Org = "MEX" Then c = "Me"
If Org = "SLB" Then c = "s"

Application.SendKeys (c), True

End Sub
Sub changeOrg()
'this is for changing orgs in Bom and routing section
Org = Cells(findCell.Row, 11).Value
slow1

Application.SendKeys ("5"), True
slow1

Dim c As String
If Org = "CNL" Then c = "cn"
If Org = "GWH" Then c = "g"
If Org = "LVG" Then c = "l"
If Org = "MEX" Then c = "Me"
If Org = "SLB" Then c = "s"

Application.SendKeys (c), True

End Sub

Sub setvars()
'this sub sets assigns the values of item variables which are used by other subs
t = findCell.Column


AlphaNum = Cells(findCell.Row, 2).Value
ItemNum = Left(AlphaNum, 5)
Cells(findCell.Row, 3) = ItemNum
Org = Cells(findCell.Row, 11).Value
Cost = Cells(findCell.Row, 12).Value

End Sub
Sub submitAlpha()
'take info from Alpha form and load it into Alpha Sheet

Set opRng = Worksheets("ALPHA PARTS WIP").Range("B2:B35")
Application.FindFormat.Clear
Application.FindFormat.Interior.ColorIndex = 2

Dim rng As Range
Set rng = ThisWorkbook.Worksheets("ALPHA PARTS WIP").Range("B2:B113")
Set findCell = rng.Find(what:="*", searchFormat:=True)
If (findCell Is Nothing) Then
    BringToFront1
    MsgBox "No Unprocessed Items"
    Exit Sub
    

End If
'iRow = Cells(Rows.Count, 2).End(xlUp).Row + 1

ThisWorkbook.Worksheets("LocArray").Unprotect Password:="1234"
 Dim ph As Worksheet
    Dim text1 As String
    Dim text2 As String
    Dim Counter As Integer
    Dim List As Variant
    Dim Item As Variant
    Dim rng1 As Range
    Dim LocRng As Range
    Set LocRng = ThisWorkbook.Worksheets("LocArray").Range("v3:v7")
    
    
    Set ph = ThisWorkbook.Sheets("ALPHA PARTS WIP")

    ThisWorkbook.Sheets("LocArray").Cells(3, 22) = AlphaForm.TextBox1.Value
    ThisWorkbook.Sheets("LocArray").Cells(4, 22) = AlphaForm.TextBox2.Value
    ThisWorkbook.Sheets("LocArray").Cells(5, 22) = AlphaForm.TextBox3.Value
    ThisWorkbook.Sheets("LocArray").Cells(6, 22) = AlphaForm.TextBox4.Value
    ThisWorkbook.Sheets("LocArray").Cells(7, 22) = AlphaForm.TextBox5.Value
    
    
    ThisWorkbook.Sheets("LocArray").Cells(3, 23) = AlphaForm.TextBox6.Value
    ThisWorkbook.Sheets("LocArray").Cells(4, 23) = AlphaForm.TextBox7.Value
    ThisWorkbook.Sheets("LocArray").Cells(5, 23) = AlphaForm.TextBox8.Value
    ThisWorkbook.Sheets("LocArray").Cells(6, 23) = AlphaForm.TextBox9.Value
    ThisWorkbook.Sheets("LocArray").Cells(7, 23) = AlphaForm.TextBox10.Value
    
    If AlphaForm.CheckBox1.Value = True Then Cells(findCell.Row, 15) = "TRUE"
' this later section is probably obsolete, check later
    
    For Each rng1 In LocRng  ' take list and create lines in display sheet
        If rng1 = "" Then GoTo Skip
        
    text1 = text1 & rng1.Value & vbCrLf
    text2 = text2 & rng1.Offset(0, 1).Value & vbCrLf

Skip:
    Next rng1
    If text1 = "" Then GoTo EndPoint
    text1 = Left(text1, Len(text1) - 2)
    text2 = Left(text2, Len(text2) - 2)
    
    
    Cells(findCell.Row, 13) = text1
    Cells(findCell.Row, 14) = text2
    
    Cells(findCell.Row, 2) = AlphaForm.TextBox11.Value
    Cells(findCell.Row, 12) = AlphaForm.TextBox12.Value
    
EndPoint:
Stop
ThisWorkbook.Worksheets("LocArray").Protect Password:="1234"
End Sub
Sub ShowAlpha_Form()
'call for Alpha form
    AlphaForm.Show
End Sub


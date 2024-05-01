Attribute VB_Name = "Module6"
Public Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As LongPtr, ByVal Y As LongPtr) As LongPtr
Public Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwmilliseconds As LongPtr)

Public Declare PtrSafe Sub mouse_event Lib "user32" (ByVal dwFlags As LongPtr, ByVal dx As Long, ByVal dy As LongPtr, ByVal cButtons As LongPtr, ByVal swextrainfo As LongPtr) '
Public Const mouseeventf_Leftdown = &H2
Public Const mouseeventf_Leftup = &H4
Public Const mouseeventf_Rightdown As Long = &H8
Public Const mouseeventf_rightup As Long = &H10
Public txt1 As String
Option Explicit
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
Sub debug2()
If Right(Cells(8, 12), 1) = vbLf Then MsgBox "yes"

End Sub

Sub removeEMPTY()

Cells(9, 11) = Replace(Cells(9, 11), " ", "")
Cells(9, 12) = Replace(Cells(9, 12), " ", "")

Cells(9, 11) = Replace(Cells(9, 11), vbLf, "")
Cells(9, 12) = Replace(Cells(9, 12), vbLf, "")

Cells(9, 11) = Replace(Cells(9, 11), Chr(13), "")
Cells(9, 12) = Replace(Cells(9, 12), Chr(13), "")
Cells(9, 11) = Replace(Cells(9, 11), Chr(9), "")
Cells(9, 12) = Replace(Cells(9, 12), Chr(9), "")
End Sub


Sub CloseWordWindows()

Dim objWord As Object

On Error Resume Next
Set objWord = GetObject(, "Word.Application")

' if no active Word is running >> exit Sub
If objWord Is Nothing Then
    Exit Sub
End If

objWord.Quit
Set objWord = Nothing

End Sub
Sub convertToWord()

Dim MyObj As Object
Dim MySource As Object
Dim file As Variant
Dim doc As Word.document
Dim path As String

file = Dir("H:\My Documents\Custom_Item_Cost_Report_271123" & "*.pdf")
path = "H:\My Documents\"
ChDir "H:\My Documents\"


        'Open method will automatically convert PDF for editing
        Set doc = Documents.Open(path & file, False)

        'Save and close document
        doc.SaveAs2 path & Replace(file, ".pdf", ".docx"), _
                    FileFormat:=wdFormatDocumentDefault
        doc.Close False
        WORDSelectDOCandCopytoA1
        
    CloseWordWindows

     'file = Dir



End Sub

Sub colormepurple()

Cells(18, 8).Interior.ColorIndex = 12
End Sub
Sub giveUserIN()
MsgBox Environ("UserName")
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

Dim RNG1 As Range
Dim RNG2 As Range
Dim findCellCol As Range
Dim MorA As String
MorA = "assm"
Dim AifType As String
AifType = "transfer"
Dim Org As String
Org = "MEX"
Dim OrgRow As Range
Dim CellCode As String
    
    
Set RNG1 = ThisWorkbook.Worksheets("LocArray").Range("i4:i12")
Set RNG2 = ThisWorkbook.Worksheets("LocArray").Range("j3:n3")

Set findCellCol = RNG1.Find(what:=AifType, MatchCase:=False)
If findCellCol Is Nothing Then
Set findCellCol = RNG1.Find(what:=MorA, MatchCase:=True)
Else
End If
Set OrgRow = RNG2.Find(what:=Org, MatchCase:=False)

'Set CellCode = Cells(OrgRow, findCellCol).Value

MsgBox Cells(findCellCol.Row, OrgRow.Column).Value
Stop



End Sub
Sub changeColor()

ThisWorkbook.Worksheets("Test").Cells(14, 9).Value = "Completed"

End Sub
Sub slow1()
'sub to cause 1 sec delay in process
DoEvents
Application.Wait (Now + TimeValue("00:00:01"))

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
    OpenClipboard (0&)
    EmptyClipboard
    CloseClipboard
End Sub

Sub activateEDGE()
On Error Resume Next
    AppActivate ("Internet explorer")
On Error GoTo 0
On Error Resume Next
    AppActivate ("http://ebs.nifconet.com:8000/OA_CGI")
On Error GoTo 0
CopyCompareCell
If (InStr(1, txt1, "What do you want to do with Custom_Item_Cost_Report")) > 0 Then
    Application.SendKeys "%a"
    slow1
    ClearClipboard
    CopyCompareCell
    If (InStr(1, txt1, "Custom_Item_Cost_Report_")) > 0 Then
     Application.SendKeys "yes"
     slow1
     Application.SendKeys "%s"
     End If
End If

End Sub

Sub copylist()

'slow2
ClickOnCorner1Window
Application.Wait (Now + TimeValue("00:00:01"))

Dim Mer As Range
'Mer = ThisWorkbook.Sheets(Sheet1).Range("h16:h281")
Dim rowI As Integer
rowI = 9
For rowI = 9 To 45

'MsgBox (copyCell)
slowMill
Application.SendKeys ("<")
slowMill
Application.SendKeys ThisWorkbook.Worksheets("KickOFF").Cells(rowI, 1).Value
slowMill
Application.SendKeys ("<, _")
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
Sub WORDSelectDOCandCopytoA1()
    CloseWordWindows
    'Create variables
    Dim Word As New Word.Application
    Dim WordDoc As New Word.document
    Dim r As Word.Range
    Dim Doc_Path As String
    Dim WB As Excel.Workbook
    Dim WB_Name As String
    Word.Visible = False
    Doc_Path = "H:\My Documents\Custom_Item_Cost_Report_271123.docx"
    Set WordDoc = Word.Documents.Open(Doc_Path)
    ' Set WordDoc = ActiveDocument

    ' Create a range to search.
    ' All of content is being search here
    Set r = WordDoc.Content

    'Find text and copy it (part that I am having trouble with)
    With r
        .Find.ClearFormatting
        With .Find
            .Text = "System Cost"
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = True
            .Execute
        End With
        r.EndOf Unit:=wdParagraph, Extend:=wdExtend
        .Copy
        ' Debug.Print r.Text
    End With


    'Open excel workbook and paste
    ThisWorkbook.Worksheets("TEST").Cells(1, 1).Select
    
    Dim ClipObj As New DataObject
    
    ClipObj.GetFromClipboard
    On Error GoTo 0

ThisWorkbook.Worksheets("TEST").Range("A1") = ClipObj.GetText(1)


    WordDoc.Close
    Word.Quit
    


End Sub
Sub WORDExtractText()

Dim cDoc As Word.document
Dim crng As Word.Range
Dim i As Long
i = 2
Dim wordapp As Object
Set wordapp = CreateObject("word.Application")
wordapp.Documents.Open "c:\bracketdata\bracket-data.docx"
wordapp.Visible = True
Set cDoc = ActiveDocument
Set crng = cDoc.Content
    With crng.Find
        .Forward = True
        .Text = "["
        .wrap = wdFindStop
        .Execute
        Do While .Found
            'Collapses a range or selection to the starting or ending position
            crng.collapse
        Word.WdCollapseDirection.wdCollapseEnd
            crng.MoveEndUntil Cset:="]"
            Cells(i, 1) = crng
            crng.collapse
        Word.WdCollapseDirection.wdCollapseEnd
        .Execute
        i = i + 1
    Loop
End With
wordapp.Quit
Set wordapp = Nothing
End Sub
Option Explicit
Sub achieID()

MsgBox Asc(Cells(8, 12))

End Sub
Sub ReadDOc()
    
    'Create variables
    'Dim Word As New Word.Application
    'Dim WordDoc As New Word.document
    Dim r As Word.Range
    Dim Doc_Path As String
    'MyWord.Visible = True
    Dim itemTag As String
    
    itemTag = WorksheetFunction.Concat(Item, " Custom Cost Sheet ", CostType, " ", WorksheetFunction.Text(Date, "mm-dd-yy"))
    Doc_Path = "H:\My Documents\" & itemTag & ".docx"
    'Set WordDoc = MyWord.Documents.Open(Doc_Path)
    ' Set WordDoc = ActiveDocument

    ' Create a range to search.
    ' All of content is being search here
    Set r = WordDoc.Content

    'Find text and copy it (part that I am having trouble with)
    With r
        .Find.ClearFormatting
        With .Find
            .Text = "System Cost"
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = True
            .Execute
        End With
        r.EndOf Unit:=wdSentence, Extend:=wdExtend
        'r.EndOf Unit:=wdStory, Extend:=wdExtend
        .Copy
        ' Debug.Print r.Text
    End With
   
    
ThisWorkbook.Worksheets("AIF").Cells(findCell.Row, 11) = Trim(Mid(Left(r, InStr(1, r, "Total") - 1), 13, 20))
ThisWorkbook.Worksheets("AIF").Cells(findCell.Row, 12) = Trim(Mid(Mid(r, InStr(1, r, "Total"), 100), 11, 20))
slow1
ThisWorkbook.Worksheets("AIF").Cells(findCell.Row, 11) = Replace(ThisWorkbook.Worksheets("AIF").Cells(findCell.Row, 11), Chr(7), "")
ThisWorkbook.Worksheets("AIF").Cells(findCell.Row, 12) = Replace(ThisWorkbook.Worksheets("AIF").Cells(findCell.Row, 12), Chr(7), "")
ThisWorkbook.Worksheets("AIF").Cells(findCell.Row, 11) = Replace(ThisWorkbook.Worksheets("AIF").Cells(findCell.Row, 11), " ", "")
ThisWorkbook.Worksheets("AIF").Cells(findCell.Row, 12) = Replace(ThisWorkbook.Worksheets("AIF").Cells(findCell.Row, 12), " ", "")
ThisWorkbook.Worksheets("AIF").Cells(findCell.Row, 11) = Replace(ThisWorkbook.Worksheets("AIF").Cells(findCell.Row, 11), vbLf, "")
ThisWorkbook.Worksheets("AIF").Cells(findCell.Row, 12) = Replace(ThisWorkbook.Worksheets("AIF").Cells(findCell.Row, 12), vbLf, "")
ThisWorkbook.Worksheets("AIF").Cells(findCell.Row, 11) = Replace(ThisWorkbook.Worksheets("AIF").Cells(findCell.Row, 11), Chr(13), "")
ThisWorkbook.Worksheets("AIF").Cells(findCell.Row, 12) = Replace(ThisWorkbook.Worksheets("AIF").Cells(findCell.Row, 12), Chr(13), "")
ThisWorkbook.Worksheets("AIF").Cells(findCell.Row, 11) = Replace(ThisWorkbook.Worksheets("AIF").Cells(findCell.Row, 11), Chr(9), "")
ThisWorkbook.Worksheets("AIF").Cells(findCell.Row, 12) = Replace(ThisWorkbook.Worksheets("AIF").Cells(findCell.Row, 12), Chr(9), "")


'ThisWorkbook.Worksheets("AIF").Cells(findCell.Row, 12) = r

slow1
'ThisWorkbook.Worksheets("TEST").Range("A1") = ClipObj.GetText(1)
'Stop
WordDoc.Close False
slow1
'Stop
'MyWord.Quit
slow1
'Stop
Kill Doc_Path
'Stop

End Sub
Sub WORDExtractTextEP()

End Sub
Dim cDoc As Word.document, nDoc As Word.document
Dim crng As Word.Range, nRng As Word.Range
Set cDoc = ActiveDocument
Set nDoc = Documents.Add
Set crng = cDoc.Content
Set nRng = nDoc.Content
crng.Find.ClearFormatting
With crng.Find
    .Forward = True
    .Text = "["
    .wrap = wdFindStop
    .Execute
    Do While .Found
        crng.collapse
Word.WdCollapseDirection.wdCollapseEnd
    crng.MoveEndUntil Cset:="]", count:=Word.wdForward

    nRng.FormattedText = crng.FormattedText
    nRng.InsertParagraphAfter
    nRng.collapse
Word.WdCollapseDirection.wdCollapseEnd
    crng.collapse
Word.WdCollapseDirection.wdCollapseEnd
    .Execute
    Loop
    End With
End Sub
Sub WORDCopyTPNumber()

    'Create variables
    Dim Word As New Word.Application
    Dim WordDoc As New Word.document
    Dim r As Word.Range
    Dim Doc_Path As String
    Dim WB As Excel.Workbook
    Dim WB_Name As String

    Doc_Path = "C:\temp\TestFind.docx"
    Set WordDoc = Word.Documents.Open(Doc_Path)
    ' Set WordDoc = ActiveDocument

    ' Create a range to search.
    ' All of content is being search here
    Set r = WordDoc.Content

    'Find text and copy it (part that I am having trouble with)
    With r
        .Find.ClearFormatting
        With .Find
            .Text = "TP[0-9]{4}"
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = True
            .Execute
        End With
        .Copy
        ' Debug.Print r.Text
    End With


    'Open excel workbook and paste
    WB_Name = Excel.Application.GetOpenFilename(",*.xlsx")
    Set WB = Workbooks.Open(WB_Name)

    WB.Sheets("Sheet1").Select
    Range("AB2").Select
    ActiveSheet.Paste
    WordDoc.Close
    Word.Quit

End Sub
Sub WORDInsertFromFilesTestEnd()


Dim wrdApp As Word.Application
Dim wrdDoc As Word.document


    Set wrdApp = New Word.Application
    Set wrdDoc = wrdApp.Documents.Open("c:\users\peter\documents\direkte 0302 1650.docm")
        
        wrdApp.Visible = True
        wrdApp.Activate
        Application.ScreenUpdating = False

wrdApp.Selection.EndKey Unit:=wdStory, Extend:=wdMove
End Sub

Private Sub WORDCommandButton3_Click()
    Set wordapp = CreateObject("word.Application")

    On Error GoTo OpenedErr

    Set doc = wordapp.Documents.Open("C:\Users\rossy\OneDrive\Work In Progress\Payroll and Billing Spreadsheet\Newest 148\Code\1.docx")
   
    wordapp.Application.WindowState = wdWindowStateNormal
    wordapp.Application.Resize Width:=400, Height:=400
    
    wordapp.Visible = True
    wordapp.Application.Activate

OpenedErr:
    ' Don´t forget to clean memory once done
    Set doc = Nothing
    Set wordapp = Nothing
End Sub
Dim appWord As Word.Application
Dim document As Word.document

Set appWord = CreateObject("Word.Application")

' the excel file and the word document are in the same folder
Set document = appWord.Documents.Open( _
ThisWorkbook.path & "\Testfile.docx")

' adding the needed data to the word file
...

' in this following line of code I tried to do the correct saving... but it opens the "save as" window - I just want to save it automatically
document.Close wdSaveChanges:=-1

'Close our instance of Microsoft Word
appWord.Quit

'Release the external variables from the memory
Set document = Nothing
Set appWord = Nothing
End Sub
Public Sub WORDWordFindAndReplaceTEST()
    Dim ws As Worksheet, msWord As Object
    Dim firstTerm As String
    Dim secondTerm As String
    Dim documentText As String
    
    
    'Dim myRange As Range
    Dim myRange As Word.Range ' just like  Timothy Rylatt said
    
    Dim startPos As Long 'Stores the starting position of firstTerm
    Dim stopPos As Long 'Stores the starting position of secondTerm based on first term's location
    Dim nextPosition As Long 'The next position to search for the firstTerm
    
    Dim d As Word.document 'use this to OP the opened document instead of ActiveDocument

    nextPosition = 1

    firstTerm = "<Tag2.1.1>"
    secondTerm = "</Tag2.1.1>"
    
    On Error Resume Next
    Set msWord = GetObject(, "Word.Application")
    Rem wrdApp should be msWord
    'If wrdApp Is Nothing Then
    If msWord Is Nothing Then
        Set msWord = CreateObject("Word.Application")
    End If
    On Error GoTo 0

    Set ws = ActiveSheet
    

    With msWord
        .Visible = True
        
        '.Documents.Open "C:\Users\user\Desktop\ReportTest\ReportDoc.rtf"
        
        Rem using this to OP the opened document instead of ActiveDocument is better
        Set d = .Documents.Open("C:\Users\user\Desktop\ReportTest\ReportDoc.rtf")
'        Set d = .Documents.Open("X:\PS Test\1.rtf") 'this for my test
        
        .Activate
        
            'Get all the document text and store it in a variable.
            'documentText = .ActiveDocument.Content
            documentText = d.Content
                        
            Rem Using Word.Range object, you have to initiate it by `Set` that object to a range in a document first
            Set myRange = d.Range
            
            'Loop documentText till you can't find any more matching "terms"
            Do Until nextPosition = 0
                startPos = InStr(nextPosition, documentText, firstTerm, vbTextCompare)
                stopPos = InStr(startPos, documentText, secondTerm, vbTextCompare)
                nextPosition = InStr(stopPos, documentText, firstTerm, vbTextCompare)
            'Loop ' Wrong place to close the loop!!
            
                'Set myRange = Nothing 'this is meanless!
                myRange.SetRange Start:=startPos, End:=stopPos 'Error thrown here
                'MsgBox .ActiveDocument.Range(startPos, stopPos) 'Successfully returns range as string
                
                'myRange.Select 'just for test to check out
                
                'With .ActiveDocument.Content.Find
                'With d.Content.Find' this will replace all text of the opened file not only the range
                With myRange.Find
                    .ClearFormatting
                    .Replacement.ClearFormatting
                    
                    .Text = "toReplace"
                    .Replacement.Text = "replacementText"
        
                    .Forward = True
                    '.Wrap = 1 'wdFindContinue' this will replace all text of the opened file not only the range
                    .wrap = wdFindStop
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                    .Execute Replace:=2 'wdReplaceAll
                End With
            
            Loop
            
        'Overrides original
        '.Quit SaveChanges:=True 'this will save all your files if `GetObject(, "Word.Application")` succeed.
        If Not d.Saved Then
            d.Close Word.wdSaveChanges
        Else
            d.Close 'when there is nothing to be replaced. However, open .rtf files in MS Word seem to be modified.
        End If
        If .Documents.count = 0 Then .Quit
        
    End With
End Sub



Sub BringToFront()
    Dim setFocus As Long
    
    ThisWorkbook.Worksheets("database").Activate
    setFocus = SetForegroundWindow(Application.hWnd)
End Sub

Sub f5()
ClickOnCornerWindow
Application.SendKeys ("{F5}")
End Sub
Sub findlastWordinColumn()

Dim lRow As Long

lRow = Cells(Rows.count, 2).End(xlUp).Row

MsgBox lRow

End Sub

Public Sub StartExeWithArgument()
    Dim strFilename As String

    strFilename = "../folder/file.pdf"

    Call Shell(strFilename, vbNormalFocus)
End Sub

Sub chromeshell()
'The path to the file, replaces spaces with the encoding "%20"
Path = Replace((filePath& "#Page=" & iPageNum), " ", "%20")

Dim wshShell, chromePath As String, shellPath As String
Set wshShell = CreateObject("WScript.Shell")
chromePath = wshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe\")

shellPath = CStr(chromePath) & " -url " & path

If Not chromePath = "" Then
    'how I first tried it:
    Shell (shellPath)

    'for testing purposes, led to the same result though:
    Shell ("""C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"" ""C:\Users\t.weinmuellner\Desktop\Evon\PDF Opening\PDFDocument.pdf#page=17""")

End If

End Sub

Sub edgeshell()
Dim oShell As Object
Dim strFilename As String

strFilename = "file:///H:\My Documents\xyz.pdf"
strFilename = Replace(strFilename, " ", "%20")
MsgBox strFilename

Set oShell = CreateObject("WScript.Shell")

'oShell.Run "notepad " & strFilename
'oShell.Run "iexplore " & strFilename
'oShell.Run "winword " & strFilename
oShell.Run "msedge " & strFilename
End Sub

Attribute VB_Name = "modCellularAutomata"
Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''My apologies for the source currently being completely undocumented.  I want to go ahead and submit an initial version for those who are interested'''
'''in Wolfram's work, and the code is clean and straight-forward enough that most can probably figure it out.                                         '''
'''                                                                                                                                                   '''
'''Thanks for trying my program,                                                                                                                      '''
'''     Neal Blair                                                                                                                                    '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BitBlt() Declarations:
Declare Function BitBlt Lib "gdi32" ( _
        ByVal hDestDC As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal hSrcDC As Long, _
        ByVal xSrc As Long, _
        ByVal ySrc As Long, _
        ByVal dwRop As Long) As Long

Public Const SRCCOPY = &HCC0020
Public Const SRCAND = &H8800C6
Public Const SRCINVERT = &H660046
Public Const SRCPAINT = &HEE0086
Public Const SRCERASE = &H4400328
Public Const WHITENESS = &HFF0062
Public Const BLACKNESS = &H42
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Global ZOOMFACTOR As Integer
Global SEED As Integer
Global ITERATIONS As Integer
Global PRINTMODE As String
Global ReplaceString As String
Global ReplaceColor(7) As Integer
Global LineString() As String

Global DisplayX As Integer
Global DisplayY As Integer


Public Function ColorToValue(ColorNumber As Integer) As Double

Select Case ColorNumber
    Case 1:
        ColorToValue = 0
    Case 2:
        ColorToValue = 1 / 3
    Case 3:
        ColorToValue = 2 / 3
    Case 4:
        ColorToValue = 1
    Case 5:
        ColorToValue = 4 / 3
    Case 6:
        ColorToValue = 5 / 3
    Case 7:
        ColorToValue = 2
End Select

End Function


Public Function ValueToColor(Value As Double) As Integer

Value = Round(Value, 2)

Select Case Value
    Case 0:
        ValueToColor = 1
    Case 0.33:
        ValueToColor = 2
    Case 0.67:
        ValueToColor = 3
    Case 1:
        ValueToColor = 4
    Case 1.33:
        ValueToColor = 5
    Case 1.67:
        ValueToColor = 6
    Case 2:
        ValueToColor = 7
End Select

End Function

Public Function GetAverageColor(TriString As String) As Integer

GetAverageColor = ValueToColor((ColorToValue(Mid(TriString, 1, 1)) + ColorToValue(Mid(TriString, 2, 1)) + ColorToValue(Mid(TriString, 3, 1))) / 3)

End Function

Public Function ProcessString(IterationNumber As Integer) As String
DoEvents

Dim WorkingString As String
Dim NextString As String
Dim FillCount As Integer

For FillCount = IterationNumber To 1 Step -1
    LineString(FillCount) = "11" & LineString(FillCount) & "11"
Next FillCount

WorkingString = LineString(IterationNumber)
NextString = ""

Dim SpotCount As Integer

For SpotCount = 1 To Len(WorkingString) - 2
    NextString = NextString & ReplaceColor(GetAverageColor(Mid(WorkingString, SpotCount, 3)))
Next SpotCount

NextString = "1" & NextString & "1"

ProcessString = NextString
End Function

Public Function PrintAutomaton(PrintType As String)
Dim IterationCount As Integer
Dim SavePath As String

SavePath = App.Path
If Right(SavePath, 1) <> "\" Then SavePath = SavePath & "\"

Select Case PrintType
    Case "Text":
        Open SavePath & "Automata-" & SEED & " Seeded-Code " & ReplaceString & "-" & ITERATIONS & " Iterations--" & Replace(Replace(Replace(Replace(Replace(DateTime.Now, "/", ""), ":", ""), " ", ""), "PM", "P"), "AM", "A") & ".txt" For Output As #1
        
        Print #1, "Seed Color: " & SEED
        Print #1, "Replace Colors: " & ReplaceColor(1) & ReplaceColor(2) & ReplaceColor(3) & ReplaceColor(4) & ReplaceColor(5) & ReplaceColor(6) & ReplaceColor(7)
        Print #1, "Iterations: " & ITERATIONS & vbCrLf

        For IterationCount = 1 To ITERATIONS
            frmMain.Caption = "Printing: " & IterationCount
            Print #1, Replace(LineString(IterationCount), "1", " ")
        Next IterationCount
        
    Case "HTML":
        Open SavePath & "Automata-" & SEED & " Seeded-Code " & ReplaceString & "-" & ITERATIONS & " Iterations--" & Replace(Replace(Replace(Replace(Replace(DateTime.Now, "/", ""), ":", ""), " ", ""), "PM", "P"), "AM", "A") & ".html" For Output As #1
        
        Print #1, "<html><body>"
        Print #1, "Seed Color: " & SEED & "<br>"
        Print #1, "Replace Colors: " & ReplaceColor(1) & ReplaceColor(2) & ReplaceColor(3) & ReplaceColor(4) & ReplaceColor(5) & ReplaceColor(6) & ReplaceColor(7) & "<br>"
        Print #1, "Iterations: " & ITERATIONS & "<br>"
        Print #1, "<font face=" & Chr(34) & "Courier New" & Chr(34) & ">"
        
        For IterationCount = 1 To ITERATIONS
            frmMain.Caption = "Printing: " & IterationCount
            Print #1, Replace(Replace(Replace(LineString(IterationCount), "1", "&nbsp;"), "4", "<span style=" & Chr(34) & "background-color: #808080" & Chr(34) & ">&nbsp;</span>"), "7", "<span style=" & Chr(34) & "background-color: #000000" & Chr(34) & ">&nbsp;</span>")
        Next IterationCount
End Select

Close #1

End Function

Public Function GenerateWindow()

Dim XCount As Integer
Dim YCount As Integer
Dim SHeight As Integer
Dim SWidth As Integer
Dim ColorValue As Integer
Dim DisplayIterations As Integer

frmDisplay.Cls

For YCount = 0 To DisplayY - 1
    For XCount = 0 To DisplayX
        ColorValue = Mid(LineString(YCount + frmDisplay.vsbDisplay.Value), XCount + frmDisplay.hsbDisplay.Value, 1)
        If ColorValue <> 1 Then BitBlt frmDisplay.hDC, XCount * ZOOMFACTOR, YCount * ZOOMFACTOR, ZOOMFACTOR, ZOOMFACTOR, frmDisplay.Color(ColorValue).hDC, 1, 1, vbSrcCopy
    Next XCount
Next YCount

frmDisplay.Refresh

End Function

Public Function MakeString(msIterations As Integer)
ReDim LineString(msIterations)
End Function

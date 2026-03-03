
Sub GlobalDeleteFirstSpace()
''''---------Ask---------------
issStartMacros = MsgBox("Run macros ""GlobalDeleteFirstSpace""? Works for a long time", vbYesNo + vbQuestion)
If issStartMacros = vbYes Then
''''-----------Timer------------
Dim StartTime As Double
Dim SecondsElapsed As Double
StartTime = Timer
''''------------------------

Dim j As Integer
For j = 1 To ActiveDocument.Paragraphs.Count

DeleteFirstSpace (j)

 Next j
 
 
''''-------Show Time-----------------
SecondsElapsed = Round(Timer - StartTime)
MsgBox "This code ran successfully in " & SecondsElapsed & " seconds", vbInformation
''''---------End Ask---------------
End If
End Sub
Sub GlobalDeleteFirstSpaceNoAsk()

''''-----------Timer------------
Dim StartTime As Double
Dim SecondsElapsed As Double
StartTime = Timer
''''------------------------

Dim j As Integer
For j = 1 To ActiveDocument.Paragraphs.Count

DeleteFirstSpace (j)

 Next j
 
 
''''-------Show Time-----------------
SecondsElapsed = Round(Timer - StartTime)
'MsgBox "This code ran successfully in " & SecondsElapsed & " seconds", vbInformation
''''---------End Ask---------------

End Sub

Private Function DeleteFirstSpace(index As Integer)

ActiveDocument.Paragraphs(index).Range.Select

char1 = Left(Selection.Paragraphs(1).Range, 1)

 If char1 = " " Then
  Selection.Paragraphs(1).Range.Characters(1).Select
 'Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
  Selection.Delete
 End If

'Selection.MoveDown Unit:=wdParagraph, Count:=1

End Function

Sub set_wdRowHeightAuto()

''''------------------------
issStartMacros = MsgBox("Start macros ""RowHeightAuto""?", vbYesNo + vbQuestion)
If issStartMacros = vbYes Then
''''------------------------

Dim mytable1 As Table
 
For Each mytable1 In ActiveDocument.Tables
mytable1.Range.Editors.Add wdEditorEveryone

mytable1.Rows.HeightRule = wdRowHeightAtLeast


        'mytable1.TopPadding = InchesToPoints(0#)
        'mytable1.BottomPadding = InchesToPoints(0#)
        'mytable1.LeftPadding = InchesToPoints(0#)
        'mytable1.RightPadding = InchesToPoints(0#)


'mytable1.Rows.HeightRule = wdRowHeightAuto
'mytable1.Rows.Height = InchesToPoints(0)
Next
ActiveDocument.SelectAllEditableRanges (wdEditorEveryone)
ActiveDocument.DeleteAllEditableRanges (wdEditorEveryone)
'Call set_hederGlobal
''''------------------------
End If
''''------------------------

End Sub

Sub set_wdRowHeightAutoNoAsk()

''''------------------------

Dim mytable1 As Table
 
For Each mytable1 In ActiveDocument.Tables
mytable1.Range.Editors.Add wdEditorEveryone

mytable1.Rows.HeightRule = wdRowHeightAtLeast

'mytable1.Rows.HeightRule = wdRowHeightAuto
'mytable1.Rows.Height = InchesToPoints(0)
Next
ActiveDocument.SelectAllEditableRanges (wdEditorEveryone)
ActiveDocument.DeleteAllEditableRanges (wdEditorEveryone)

''''------------------------


End Sub

Sub set_wdAutoFitWindow()
''''------------------------
issStartMacros = MsgBox("Start macros ""AutoFitWindow""?", vbYesNo + vbQuestion)
If issStartMacros = vbYes Then
''''------------------------

Dim mytable As Table
For Each mytable In ActiveDocument.Tables
mytable.Range.Editors.Add wdEditorEveryone
mytable.AutoFitBehavior (wdAutoFitWindow) '
mytable.Rows.WrapAroundText = False
With mytable
       
        .AllowPageBreaks = False '
        .AllowAutoFit = False '
    End With



Next
ActiveDocument.SelectAllEditableRanges (wdEditorEveryone)
ActiveDocument.DeleteAllEditableRanges (wdEditorEveryone)

''''------------------------
End If
''''------------------------
End Sub

Sub set_wdAutoFitWindowNoAsk()
''''------------------------


Dim mytable As Table
For Each mytable In ActiveDocument.Tables
mytable.Range.Editors.Add wdEditorEveryone
mytable.AutoFitBehavior (wdAutoFitWindow)

With mytable
       
        .AllowPageBreaks = False
        .AllowAutoFit = False
    End With



Next
ActiveDocument.SelectAllEditableRanges (wdEditorEveryone)
ActiveDocument.DeleteAllEditableRanges (wdEditorEveryone)

''''------------------------

End Sub


Sub conversion()
    '------------------------------------------------
    ' Preparation for conversion _SO
    '------------------------------------------------
    
    Dim issStartMacros As Integer
    issStartMacros = MsgBox("Start macros ""conversion""?", vbYesNo + vbQuestion)
    If issStartMacros <> vbYes Then Exit Sub
    
    ' Початок запису для Undo
    Application.UndoRecord.StartCustomRecord "Conversion Macro"

    On Error GoTo column

    'Set one column in the whole document
    Selection.WholeStory
    If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        ActiveWindow.Panes(2).Close
    End If
    If ActiveWindow.ActivePane.View.Type <> wdPrintView Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    End If
    With Selection.PageSetup.TextColumns
        .SetCount NumColumns:=1
        '.EvenlySpaced = True
        '.LineBetween = False
    End With
column:

    ' Remove section break
    On Error GoTo section_break
    Selection.WholeStory
    With Selection.Find
        .ClearFormatting
        .Execute findText:="^b"
        .Replacement.Text = "^p"
    End With
    
    If Selection.Find.Found Then
        Dim Result As Integer
        Result = MsgBox("Remove section break?", vbYesNo + vbQuestion)
        If Result = vbYes Then Call removeSectionBreak
    End If
section_break:

    ' Remove column break
    Selection.WholeStory
    With Selection.Find
        .ClearFormatting
        .Execute findText:="^n"
        .Replacement.Text = "^p"
    End With
    If Selection.Find.Found Then Call replaceCycle

    ' Remove double spaces
    Selection.WholeStory
    With Selection.Find
        .ClearFormatting
        .Execute findText:="  "
        .Replacement.Text = " "
    End With
    If Selection.Find.Found Then Call replaceCycle

    ' Remove space before dot
    Selection.WholeStory
    With Selection.Find
        .ClearFormatting
        .Execute findText:=" ."
        .Replacement.Text = "."
    End With
    If Selection.Find.Found Then Call replaceCycle

    ' Remove space before comma
    Selection.WholeStory
    With Selection.Find
        .ClearFormatting
        .Execute findText:=" ,"
        .Replacement.Text = ","
    End With
    If Selection.Find.Found Then Call replaceCycle

    ' Set Font Spacing, Position, Scaling
    Selection.WholeStory
    With Selection.Font
        .NameFarEast = ""
        .NameAscii = ""
        .NameOther = ""
        .Name = ""
        .Spacing = 0
        .Scaling = 100
        .Position = 0
    End With

    ' Set Line Spacing = Single
    Selection.WholeStory
    With Selection.ParagraphFormat
        .LineSpacingRule = wdLineSpaceSingle
    End With

    ' Завершення запису для Undo
    Application.UndoRecord.EndCustomRecord

End Sub

'---------------------------------------
Private Function replaceCycle()
    Do While Selection.Find.Found
        Selection.WholeStory
        Selection.Find.Execute Replace:=wdReplaceAll
    Loop
End Function

Private Function Replace(x As String, y As String)
    With Selection.Find
        .Text = x
        .Replacement.Text = y
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.Execute Replace:=wdReplaceAll
End Function

Private Function removeSectionBreak()
    Selection.WholeStory
    Dim rg As Range
    Set rg = ActiveDocument.Range
    With rg.Find
        .Text = "^b"
        .Wrap = wdFindStop
        While .Execute
            rg.Delete
            rg.InsertBreak Type:=wdPageBreak
            rg.Collapse wdCollapseEnd
        Wend
    End With

    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
End Function



Sub conversionNoAsk()
'
' preporation for conversion _SO
''''------------------------

''''------------------------

'Result4 = MsgBox("Set one column in the whole documents?", vbYesNo + vbQuestion)

On Error GoTo column

'If Result4 = vbYes Then

Selection.WholeStory
    If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        ActiveWindow.Panes(2).Close
    End If
    If ActiveWindow.ActivePane.View.Type <> wdPrintView Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    End If
    With Selection.PageSetup.TextColumns
        .SetCount NumColumns:=1
       ' .EvenlySpaced = True
      ' .LineBetween = False
    End With
'Else:
'End If
column:



'Remove section break
On Error GoTo section_break

Selection.WholeStory
    With Selection.Find
    .ClearFormatting
    .Execute findText:="^b"
    .Replacement.Text = "^m"
    End With
    
    If Selection.Find.found = True Then
       
        Call removeSectionBreak
       
    Else:
    End If

section_break:

'Remove column break?
 Selection.WholeStory
    With Selection.Find
    .ClearFormatting
    .Execute findText:="^n"
    .Replacement.Text = "^p"
    End With
  
  
    If Selection.Find.found = True Then
    '    Result = MsgBox("Remove column break?", vbYesNo + vbQuestion)
    '    If Result = vbYes Then
        Call replaceCycle
    '  Else:
    '  End If
    Else:
    End If
    


'Remove double space
 Selection.WholeStory
    With Selection.Find
    .ClearFormatting
    .Execute findText:="  "
    .Replacement.Text = " "
    End With
  
    If Selection.Find.found = True Then
        Call replaceCycle
    Else:
    End If

'Remove space before dot

    Selection.WholeStory
    With Selection.Find
    .ClearFormatting
    .Execute findText:=" ."
    .Replacement.Text = "."
    End With
  
    If Selection.Find.found = True Then
        Call replaceCycle
    Else:
    End If
   
'Remove space before comma?
    Selection.WholeStory
    With Selection.Find
    .ClearFormatting
    .Execute findText:=" ,"
    .Replacement.Text = ","
    End With
  
    If Selection.Find.found = True Then
        Call replaceCycle
    Else:
    End If
      
'Set Spacing,Position = 0 and Scaling = 100",vbYesNo + vbQuestion)
    Selection.WholeStory
    With Selection.Font
        .NameFarEast = ""
        .NameAscii = ""
        .NameOther = ""
        .Name = ""
        .Spacing = 0
        .Scaling = 100
        .Position = 0
    End With
    

'Set LineSpace = Single
'Result33 = MsgBox("Set LineSpace = Single", vbYesNo + vbQuestion)
'If Result3 = vbYes Then

    Selection.WholeStory
    With Selection.ParagraphFormat
    
        .LineSpacingRule = wdLineSpaceSingle
               
    End With
    ''''
    Call GlobalDeleteFirstSpaceNoAsk
' Else:
'End If

''''------------------------

End Sub

Sub PasteValues()
'
' PasteValues Macro
' Pastes Values only (equivalent to Ctrl+Alt+v , v, Enter).
'
' This isn't particularly robust, but it gets the job done most of the time.
'
' Keyboard Shortcut: Ctrl+Shift+V
'
    On Error GoTo ErrHandler:
    ' Catches error trying to paste otside clipboard
    ' content as xlPasteValues
    
    ' Paste as values from within Excel
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Exit Sub
    
ErrHandler:
    ' If the clipboard contents are from another program, this runs
    If Err.Number = 1004 Then
        ActiveSheet.PasteSpecial Format:="Text", Link:=False, DisplayAsIcon:= _
        False
        Exit Sub
    End If
    
End Sub
Sub append()
Debug.Print Worksheets("NewHires").Cells(Rows.Count, 1).End(xlUp).Row
   
    nhdeleter
    nhprevset = 0
    If Worksheets("NewHires").Cells(Rows.Count, 1).End(xlUp).Row > 2 Then
        nhprevset = Worksheets("NewHires").Cells(Rows.Count, 1).End(xlUp).Row
    End If
    
'NOTING YET{
    tdeleter
    termprevset = 0
    If Worksheets("Terms").Cells(Rows.Count, 1).End(xlUp).Row > 2 Then
        termprevset = Worksheets("Terms").Cells(Rows.Count, 1).End(xlUp).Row
    End If
    
    odeleter
    oprevset = 0
    If Worksheets("Other").Cells(Rows.Count, 1).End(xlUp).Row > 2 Then
        oprevset = Worksheets("Other").Cells(Rows.Count, 1).End(xlUp).Row
    End If
'}NOTING YET END
  
    trow = Worksheets(2).Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To trow
        class = Worksheets(2).Range("D" & i)
        If class = "New Hire" Then
        
            trownh = Worksheets("NewHires").Cells(Rows.Count, 1).End(xlUp).Row
            If (trownh = 2) Then
                If IsEmpty(Worksheets(2).Range("B" & i).Value) = True Then
                    sName = Worksheets(2).Range("A2").Value
                    NARRAY = Split(sName, " ")
                    Worksheets("NewHires").Range("C3").Value = NARRAY(0)
                    Worksheets("NewHires").Range("B3").Value = NARRAY(1)
                Else
                    Worksheets("NewHires").Range("B3").Value = Worksheets(2).Range("A" & i).Value
                    Worksheets("NewHires").Range("C3").Value = Worksheets(2).Range("B" & i).Value
                End If
                
                Worksheets("NewHires").Range("A3").Value = Worksheets(2).Range("C" & i).Value
                Worksheets("NewHires").Range("D3").Value = Worksheets(2).Range("F" & i).Value
                Worksheets("NewHires").Range("D5").Value = DateSerial(Year(Worksheets(2).Range("C" & i).Value), Month(Worksheets(2).Range("C" & i).Value) + 1, 1)
                Worksheets("NewHires").Range("E3").Value = Worksheets(2).Range("AE" & i).Value
                Worksheets("NewHires").Range("F3").Value = Worksheets(2).Range("AF" & i).Value
                nhprevset = 3
            Else
                trownh = Worksheets("NewHires").Cells(Rows.Count, 1).End(xlUp).Row
                nhmover (trownh - 1)
                nhprevset = nhprevset + 4
                
                If IsEmpty(Worksheets(2).Range("B" & i).Value) = True Then
                    sName = Worksheets(2).Range("A" & i).Value
                    NARRAY = Split(sName, " ")
                    Worksheets("NewHires").Range("C" & nhprevset).Value = NARRAY(0)
                    Worksheets("NewHires").Range("B" & nhprevset).Value = NARRAY(1)
                Else
                    Worksheets("NewHires").Range("B" & nhprevset).Value = Worksheets(2).Range("A" & i).Value
                    Worksheets("NewHires").Range("C" & nhprevset).Value = Worksheets(2).Range("B" & i).Value
                End If
                
                Worksheets("NewHires").Range("A" & nhprevset).Value = Worksheets(2).Range("C" & i).Value
                Worksheets("NewHires").Range("D" & nhprevset).Value = Worksheets(2).Range("F" & i).Value
                Worksheets("NewHires").Range("D" & nhprevset + 2).Value = DateSerial(Year(Worksheets(2).Range("C" & i).Value), Month(Worksheets(2).Range("C" & i).Value) + 1, 1)
                Worksheets("NewHires").Range("E" & nhprevset).Value = Worksheets(2).Range("AE" & i).Value
                Worksheets("NewHires").Range("F" & nhprevset).Value = Worksheets(2).Range("AF" & i).Value
             End If
        End If
        
'------------------------------------------------------------------------------------------------------------------------------------
        If class = "Termination" Then
            
            trowte = Worksheets("Terms").Cells(Rows.Count, 1).End(xlUp).Row
            If (trowte = 2) Then
                If IsEmpty(Worksheets(2).Range("B" & i).Value) = True Then
                    sName = Worksheets(2).Range("A" & i).Value
                    NARRAY = Split(sName, " ")
                    Worksheets("Terms").Range("C3").Value = NARRAY(0)
                    Worksheets("Terms").Range("B3").Value = NARRAY(1)
                Else
                    Worksheets("Terms").Range("B3").Value = Worksheets(2).Range("A" & i).Value
                    Worksheets("Terms").Range("C3").Value = Worksheets(2).Range("B" & i).Value
                End If
                
                Worksheets("Terms").Range("A3").Value = Worksheets(2).Range("C" & i).Value
                Worksheets("Terms").Range("D3").Value = Worksheets(2).Range("F" & i).Value
                Worksheets("Terms").Range("E3").Value = Worksheets(2).Range("AF" & i).Value
                Worksheets("Terms").Range("F3").Value = Worksheets(2).Range("Z" & i).Value
                termprevset = 3
            Else
                trowte = Worksheets("Terms").Cells(Rows.Count, 1).End(xlUp).Row
                tmover (trowte - 1)
                termprevset = termprevset + 2
                
                If IsEmpty(Worksheets(2).Range("B" & i).Value) = True Then
                    sName = Worksheets(2).Range("A" & i).Value
                    NARRAY = Split(sName, " ")
                    Worksheets("Terms").Range("C" & termprevset).Value = NARRAY(0)
                    Worksheets("Terms").Range("B" & termprevset).Value = NARRAY(1)
                Else
                    Worksheets("Terms").Range("B" & termprevset).Value = Worksheets(2).Range("A" & i).Value
                    Worksheets("Terms").Range("C" & termprevset).Value = Worksheets(2).Range("B" & i).Value
                End If
                
                Worksheets("Terms").Range("A" & termprevset).Value = Worksheets(2).Range("C" & i).Value
                Worksheets("Terms").Range("D" & termprevset).Value = Worksheets(2).Range("F" & i).Value
                Worksheets("Terms").Range("E" & termprevset).Value = Worksheets(2).Range("AF" & i).Value
                Worksheets("Terms").Range("F" & termprevset).Value = Worksheets(2).Range("Z" & i).Value
            End If
        End If
'-------------------------------------------------------------------------------------------------------------------------------------
        If class = "Other" Then
            trowo = Worksheets("Other").Cells(Rows.Count, 1).End(xlUp).Row
            If (trowo = 2) Then
                
                If IsEmpty(Worksheets(2).Range("B" & i).Value) = True Then
                    sName = Worksheets(2).Range("A2").Value
                    NARRAY = Split(sName, " ")
                    Worksheets("Other").Range("C3").Value = NARRAY(0)
                    Worksheets("Other").Range("B3").Value = NARRAY(1)
                Else
                    Worksheets("Other").Range("B3").Value = Worksheets(2).Range("A" & i).Value
                    Worksheets("Other").Range("C3").Value = Worksheets(2).Range("B" & i).Value
                End If
                
                Worksheets("Other").Range("A3").Value = Worksheets(2).Range("C" & i).Value
                Worksheets("Other").Range("D3").Value = Worksheets(2).Range("F" & i).Value
                Worksheets("Other").Range("E3").Value = Worksheets(2).Range("AF" & i).Value
                Worksheets("Other").Range("G3").Value = Worksheets(2).Range("AI" & i).Value
                oprevset = 3
            Else
                trowo = Worksheets("Other").Cells(Rows.Count, 1).End(xlUp).Row
                omover (trowo - 1)
                oprevset = oprevset + 2
                
                If IsEmpty(Worksheets(2).Range("B" & i).Value) = True Then
                    sName = Worksheets(2).Range("A" & i).Value
                    NARRAY = Split(sName, " ")
                    Worksheets("Other").Range("C" & oprevset).Value = NARRAY(0)
                    Worksheets("Other").Range("B" & oprevset).Value = NARRAY(1)
                Else
                    Worksheets("Other").Range("B" & oprevset).Value = Worksheets(2).Range("A" & i).Value
                    Worksheets("Other").Range("C" & oprevset).Value = Worksheets(2).Range("B" & i).Value
                End If
                
                Worksheets("Other").Range("A" & oprevset).Value = Worksheets(2).Range("C" & i).Value
                Worksheets("Other").Range("D" & oprevset).Value = Worksheets(2).Range("F" & i).Value
                Worksheets("Other").Range("E" & oprevset).Value = Worksheets(2).Range("AF" & i).Value
                Worksheets("Other").Range("G" & oprevset).Value = Worksheets(2).Range("AI" & i).Value
            End If
        End If
    Next
    
   
End Sub

Sub nhmover(trow)
'
' Macro3 Macro
'

'
    Worksheets("NewHires").Activate
    Rows("2:5").Select
    Application.CutCopyMode = False
    Selection.Copy
    Rows(trow + 4 & ":" & trow + 7).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("A" & trow + 4).Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Effective Date"
    Range("B" & trow + 4).Select
    ActiveCell.FormulaR1C1 = "Last Name"
    Range("C" & trow + 4).Select
    ActiveCell.FormulaR1C1 = "First Name"
    Range("D" & trow + 4).Select
    ActiveCell.FormulaR1C1 = "FLSA"
    Range("D" & trow + 6).Select
    ActiveCell.FormulaR1C1 = "Effective on:"
    Range("E" & trow + 6).Select
    ActiveCell.FormulaR1C1 = "Benefit"
    Range("E" & trow + 4).Select
    ActiveCell.FormulaR1C1 = "457(b) Election"
    Range("F" & trow + 4).Select
    ActiveCell.FormulaR1C1 = "HR Notes"
    Range("I" & trow + 6).Select
    ActiveCell.FormulaR1C1 = "HR Notes"
    Range("A" & trow + 4).Select
End Sub
Sub tmover(te)
'
' tmover Macro
'

'
    Worksheets("Terms").Activate
    Rows(te & ":" & te + 1).Select
    Selection.Copy
    Rows(te + 2 & ":" & te + 3).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("A" & te + 2).Select
    Application.CutCopyMode = False
    Range("A" & te).Select
    Selection.Copy
    Range("A" & te + 2).Select
    ActiveSheet.Paste
    Range("B" & te).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("B" & te + 2).Select
    ActiveSheet.Paste
    Range("C" & te).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C" & te + 2).Select
    ActiveSheet.Paste
    Range("D" & te).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("D" & te + 2).Select
    ActiveSheet.Paste
    Range("E" & te).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("E" & te + 2).Select
    ActiveSheet.Paste
    Range("F" & te).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("F" & te + 2).Select
    ActiveSheet.Paste
    Range("G" & te).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("G" & te + 2).Select
    ActiveSheet.Paste
    Range("H" & te).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H" & te + 2).Select
    ActiveSheet.Paste
    Range("I" & te).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("I" & te + 2).Select
    ActiveSheet.Paste
    Range("A" & te + 2).Select
    Application.CutCopyMode = False
End Sub
Sub omover(te)
'
' tmover Macro
'

'
    Worksheets("Other").Activate
    Rows(te & ":" & te + 1).Select
    Selection.Copy
    Rows(te + 2 & ":" & te + 3).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("A" & te + 2).Select
    Application.CutCopyMode = False
    Range("A" & te).Select
    Selection.Copy
    Range("A" & te + 2).Select
    ActiveSheet.Paste
    Range("B" & te).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("B" & te + 2).Select
    ActiveSheet.Paste
    Range("C" & te).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C" & te + 2).Select
    ActiveSheet.Paste
    Range("D" & te).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("D" & te + 2).Select
    ActiveSheet.Paste
    Range("E" & te).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("E" & te + 2).Select
    ActiveSheet.Paste
    Range("F" & te).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("F" & te + 2).Select
    ActiveSheet.Paste
    Range("G" & te).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("G" & te + 2).Select
    ActiveSheet.Paste
    Range("H" & te).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H" & te + 2).Select
    ActiveSheet.Paste
    Range("I" & te).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("I" & te + 2).Select
    ActiveSheet.Paste
    Range("A" & te + 2).Select
    Application.CutCopyMode = False
End Sub

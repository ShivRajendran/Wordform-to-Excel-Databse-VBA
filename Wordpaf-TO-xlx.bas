Sub paf()
    Worksheets(2).Activate
    
    Dim objWdApp As Object
    Dim objWdDoc As Object
    Set objWdApp = CreateObject("Word.Application")
    
    MsgBox "Select the PAF you want me to add to the Database", vbOKOnly
'Display a Dialog Box that allows to select a single file.
'The path for the file picked will be stored in fullpath variable
  With Application.FileDialog(msoFileDialogFilePicker)
        'Makes sure the user can select only one file
        .AllowMultiSelect = False
        'Filter to just the following types of files to narrow down selection options
        .Filters.Add "PAF Files", "*.docx", 1
        'Show the dialog box
        .Show
        
        'Store in fullpath variable
        fullpath = .SelectedItems.Item(1)
        
    End With
    
    
    Set objWdDoc = objWdApp.Documents.Open(Filename:=fullpath)
    Debug.Print objWdApp.ActiveDocument.FullName
    
    trow = Worksheets(2).Cells(Rows.Count, 1).End(xlUp).Row + 1
    objWdApp.Visible = False
    
    ActiveSheet.Range("A" & trow) = objWdDoc.textbox1.Value
    ActiveSheet.Range("B" & trow) = objWdDoc.textbox17.Value
    ActiveSheet.Range("C" & trow) = objWdDoc.ContentControls(1).Range.Text
    
    'TYPE OF ACTION
    Dim ar1 As Variant
    Dim A1 As Variant
    ar1 = Array(objWdDoc.CheckBox1.Value, objWdDoc.CheckBox2.Value, objWdDoc.CheckBox4.Value, objWdDoc.CheckBox5.Value, objWdDoc.CheckBox6.Value)
    A1 = Array("New Hire", "Termination", "Employment Changes", "Other", "401(a)/FICA Change")
    For i = 0 To 4
        If ar1(i) = True Then
            ActiveSheet.Range("D" & trow) = A1(i)
        End If
    Next
    
    
    '(1)COMPANY & (2)FLSA
    Dim ar2 As Variant
    Dim A2 As Variant
    ar2 = Array(objWdDoc.CheckBox7.Value, objWdDoc.CheckBox8.Value, objWdDoc.CheckBox9.Value, objWdDoc.CheckBox10.Value)
    A2 = Array("EDC", "OCA", "Exempt", "Non-Exempt")
    For i = 0 To 1
        If ar2(i) = True Then
            ActiveSheet.Range("E" & trow) = A2(i)
        End If
    Next
    For i = 2 To 3
        If ar2(i) = True Then
            ActiveSheet.Range("F" & trow) = A2(i)
        End If
    Next
    
    '(3)EMPLOYMENT TYPE
    Dim ar3 As Variant
    Dim A3 As Variant
    ar3 = Array(objWdDoc.CheckBox11.Value, objWdDoc.CheckBox12.Value, objWdDoc.CheckBox13.Value, objWdDoc.CheckBox14.Value, objWdDoc.CheckBox15.Value)
    A3 = Array("FT-70", "P1-21", "FT-80", "PT-28", "Intern")
    For i = 0 To 4
        If ar3(i) = True Then
            ActiveSheet.Range("G" & trow) = A3(i)
        End If
    Next
    
    '(4)PAY GRADE
    ActiveSheet.Range("H" & trow) = objWdDoc.textbox2.Value
    
    '(5)SALARY
    If objWdDoc.checkbox16.Value = True Then
        ActiveSheet.Range("I" & trow) = "Annual: " & objWdDoc.textbox3.Value
    ElseIf objWdDoc.checkbox17.Value = True Then
        ActiveSheet.Range("I" & trow) = "Hourly: " & objWdDoc.textbox3.Value
    End If
    
    
    
    '(6)MANAGEMENT POS?
    If objWdDoc.checkbox18.Value = True Then
        ActiveSheet.Range("J" & trow) = "YES"
    ElseIf objWdDoc.checkbox19.Value = True Then
        ActiveSheet.Range("J" & trow) = "NO"
    End If
    

    
    '(7)JOB TITLE
    ActiveSheet.Range("K" & trow) = objWdDoc.textbox4.Value

    
    '(8)DEPARTMENT
    ActiveSheet.Range("L" & trow) = objWdDoc.ContentControls(2).Range.Text
    If ActiveSheet.Range("L" & trow) = "Choose a department" Then ActiveSheet.Range("Y" & trow).ClearContents
    
    '(9)DIVISION
    ActiveSheet.Range("M" & trow) = objWdDoc.ContentControls(3).Range.Text
    If ActiveSheet.Range("M" & trow) = "Choose a division" Then ActiveSheet.Range("Y" & trow).ClearContents
    
    '(10)SUPERVISOR
    ActiveSheet.Range("N" & trow) = objWdDoc.textbox5.Value


    '(11)EEOC JOB CLASSIFICATION
    ActiveSheet.Range("O" & trow) = objWdDoc.ContentControls(4).Range.Text
    If ActiveSheet.Range("O" & trow) = "Choose an item" Then ActiveSheet.Range("Y" & trow).ClearContents

    '(12)EMPLOYMENT REQS
    Dim ar4 As Variant
    Dim A4 As Variant
    ar4 = Array(objWdDoc.CheckBox20.Value, objWdDoc.CheckBox21.Value, objWdDoc.CheckBox22.Value)
    A4 = Array("DOI Eligible", "COIB Eligible", "Neither")
    For i = 0 To 2
        If ar4(i) = True Then
            ActiveSheet.Range("P" & trow) = A4(i)
        End If
    Next


    '(13)PRIOR CITY SERVICE IN 5 YEARS
    Dim ar5 As Variant
    Dim A5 As Variant
    ar5 = Array(objWdDoc.CheckBox23.Value, objWdDoc.CheckBox25.Value)
    A5 = Array("YES", "NO")
    For i = 0 To 1
        If ar5(i) = True Then
            If i = 0 Then
                If objWdDoc.CheckBox24.Value = True Then
                    ActiveSheet.Range("Q" & trow) = A5(i) & ": " & objWdDoc.textbox6.Value & " verified by HRBP"
                Else
                    ActiveSheet.Range("Q" & trow) = A5(i) & ": " & objWdDoc.textbox6.Value & " not verified by HRBP"
                End If
            Else
                ActiveSheet.Range("Q" & trow) = A5(i)
            End If
        End If
    Next

  

    '(14)REFERRED BY NYCEDC EMPLOYEE
    Dim ar6 As Variant
    Dim A6 As Variant
    ar6 = Array(objWdDoc.CheckBox401.Value, objWdDoc.CheckBox411.Value)
    A6 = Array("NO", "YES")
    For i = 0 To 1
        If ar6(i) = True Then
            If i = 1 Then
                    ActiveSheet.Range("R" & trow) = A6(i) & ", by: " & objWdDoc.textbox151.Value & " , bonus amount= " & objWdDoc.ContentControls(5).Range.Text & " to be paid on " & objWdDoc.ContentControls(6).Range.Text
            Else
                ActiveSheet.Range("R" & trow) = A6(i)
            End If
        End If
    Next


    
    
    '(15)Pre-approved Vacation
    Dim ar8 As Variant
    Dim A8 As Variant
    ar8 = Array(objWdDoc.CheckBox39.Value, objWdDoc.CheckBox44.Value)
    A8 = Array("YES see dates approved: ", "NO")
    For i = 0 To 1
        If ar8(i) = True Then
            If i = 0 Then
                    ActiveSheet.Range("S" & trow) = A8(i) & objWdDoc.textbox14.Value
            Else
                ActiveSheet.Range("S" & trow) = A8(i)
            End If
        End If
    Next
    
    
    
    '(16)Relatives at NYCEDC
    Dim ar9 As Variant
    Dim A9 As Variant
    ar9 = Array(objWdDoc.CheckBox391.Value, objWdDoc.CheckBox441.Value)
    A9 = Array("YES see names below: ", "NO")
    For i = 0 To 1
        If ar9(i) = True Then
            If i = 0 Then
                    ActiveSheet.Range("T" & trow) = A9(i) & objWdDoc.textbox141.Value
            Else
                ActiveSheet.Range("T" & trow) = A9(i)
            End If
        End If
    Next
    
    
    

    '(17)NYC Residency Reqs Met
    Dim ar7 As Variant
    Dim A7 As Variant
    ar7 = Array(objWdDoc.CheckBox42.Value, objWdDoc.CheckBox421.Value)
    A7 = Array("YES", "NO")
    For i = 0 To 1
        If ar7(i) = True Then
            If i = 1 Then
                    ActiveSheet.Range("U" & trow) = A7(i) & ", HRBP follow up on: " & objWdDoc.ContentControls(7).Range.Text
            Else
                ActiveSheet.Range("U" & trow) = A7(i)
            End If
        End If
    Next
    
   
    

    '(18)Email Address
    ActiveSheet.Range("V" & trow) = objWdDoc.textbox7.Value


    '(19)Home Address change
    If objWdDoc.checkbox311.Value = True Then
        ActiveSheet.Range("W" & trow) = "No"
    ElseIf objWdDoc.checkbox312.Value = True Then
        ActiveSheet.Range("W" & trow) = "Yes, see it in notes"
    End If
    
  
    
    '(20)Termination Type
    If objWdDoc.checkbox26.Value = True Then
        ActiveSheet.Range("X" & trow) = "Voluntary"
    ElseIf objWdDoc.checkbox27.Value = True Then
        ActiveSheet.Range("X" & trow) = "Involuntary"
    End If
    
    '(21)Termination Reason
    If ActiveSheet.Range("D" & trow) = A1(1) Then
        ActiveSheet.Range("Y" & trow) = objWdDoc.ContentControls(8).Range.Text
        If ActiveSheet.Range("Y" & trow) = "Choose a Reason" Then ActiveSheet.Range("Y" & trow).ClearContents
    End If
    
    '(22)Severance Agreement
    Dim ar10 As Variant
    Dim A10 As Variant
    ar10 = Array(objWdDoc.CheckBox28.Value, objWdDoc.CheckBox29.Value)
    A10 = Array("Not Offered", "Offered")
    For i = 0 To 1
        If ar10(i) = True Then
            If i = 1 Then
                If objWdDoc.CheckBox30.Value = True Then
                    ActiveSheet.Range("Z" & trow) = A10(i) & ", Received and signed for: " & objWdDoc.textbox8.Value
                Else
                    ActiveSheet.Range("Z" & trow) = A10(i) & " not yet signed"
                End If
            Else
                ActiveSheet.Range("Z" & trow) = A10(i)
            End If
        End If
    Next
    
    

    '(23)Benefit Extension
    Dim ar11 As Variant
    Dim A11 As Variant
    ar11 = Array(objWdDoc.CheckBox31.Value, objWdDoc.CheckBox32.Value)
    A11 = Array("Not Offered", "Offered")
    For i = 0 To 1
        If ar11(i) = True Then
            If i = 1 Then
                If objWdDoc.CheckBox321.Value = True Then
                    ActiveSheet.Range("AA" & trow) = A11(i) & ", Confirmed and extended through: " & objWdDoc.ContentControls(9).Range.Text
                Else
                    ActiveSheet.Range("AA" & trow) = A11(i) & " not yet confirmed and extended"
                End If
            Else
                ActiveSheet.Range("AA" & trow) = A11(i)
            End If
        End If
    Next
    
   

    '(24)Does Employee Have Direct Reports
    If objWdDoc.checkbox33.Value = True Then
        ActiveSheet.Range("AB" & trow) = "NO"
    ElseIf objWdDoc.checkbox34.Value = True Then
        ActiveSheet.Range("AB" & trow) = "YES"
    End If
    
    
    '(25)Current ADP profile
    If objWdDoc.checkbox35.Value = True Then
        ActiveSheet.Range("AC" & trow) = "Paying into FICA"
    ElseIf objWdDoc.checkbox35.Value = True Then
        ActiveSheet.Range("AC" & trow) = "Not Paying into FICA"
    End If
    
    '(26)401a Contribution
    ActiveSheet.Range("AD" & trow) = objWdDoc.textbox9.Value
    
    '(27)457b Retirement and Savings
    If objWdDoc.checkbox37.Value = True And objWdDoc.checkbox38.Value = True Then
        ActiveSheet.Range("AE" & trow) = "Pre Tax: " & objWdDoc.textbox16.Value & " and Post Tax: " & objWdDoc.textbox10.Value
    ElseIf objWdDoc.checkbox37.Value = True Then
        ActiveSheet.Range("AE" & trow) = "Pre Tax: " & objWdDoc.textbox16.Value
    ElseIf objWdDoc.checkbox38.Value = True Then
        ActiveSheet.Range("AE" & trow) = "Post Tax: " & objWdDoc.textbox10.Value
    End If
    '(HR NOTES)
    ActiveSheet.Range("AF" & trow) = objWdDoc.textbox11.Value
    
    'PREPARED BY
     ActiveSheet.Range("AG" & trow) = objWdDoc.textbox12.Value & " on " & objWdDoc.ContentControls(10).Range.Text
     
    'APPROVED BY
    If objWdDoc.textbox13.Value <> "" Then
        ActiveSheet.Range("AH" & trow) = objWdDoc.textbox13.Value & " on " & objWdDoc.ContentControls(11).Range.Text
    Else
        ActiveSheet.Range("AH" & trow).Interior.Color = vbRed
        ActiveSheet.Range("AH" & trow).Value = "Not yet Approved!"
    End If
    
    'code for other
    
    ActiveSheet.Range("A2").Select
    If ActiveSheet.Range("D" & trow) = A1(4) Then
        cods = InputBox("I noticed this is an OTHER/SPECIAL change please input The type of change/Payroll action")
        ActiveSheet.Range("AI" & trow) = cods
        ActiveSheet.Range("A" & trow).Select
    End If
 
    
    objWdApp.Quit
    
    Set objWdDoc = Nothing
    Set objWdApp = Nothing
    Worksheets(2).Rows("2:" & Cells.Find(What:="*", After:=Range("A1"), SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row).RowHeight = 50
End Sub

Attribute VB_Name = "Module4"
Private Sub Clear_Click()
    Range("A4:DZ120000").Clear
End Sub


Private Sub Copycells_Click()
'I am Copying the Data from Sheet1 (Source) to Sheet4 (Destination)
    Sheets("Sheet1").Range("Q5:DZ120000").Copy
    Sheets("Sheet4").Range("K5").PasteSpecial Paste:=xlPasteAllMergingConditionalFormats
    'counts
    Sheets("Sheet1").Range("G2:H3").Copy
    Sheets("Sheet4").Range("E2").PasteSpecial xlPasteValues
    'divisions
    Sheets("Sheet1").Range("K5:O120000").Copy
    Sheets("Sheet4").Range("E5").PasteSpecial xlPasteValues
    
    
    
End Sub

Private Sub InsertRows_Click()

    Rows(4).Select
    Selection.Copy
    For i = 17 To 36 Step 13
        Rows(i).Select
        Selection.Insert Shift:=xlDown
        Selection.Copy
    Next i
    
   'One more loop is required i guess
   'For k = 1 To k = 12
    'For j = 18 To 37 Step 13
        'Rows(j).Select
        'Selection.Insert Shift:=xlDown
        'Selection.Copy
   ' Next j
    'Next k
    
        

    'For l = 17 To 36 Step 13

    'Range("D4:D17").Copy
   ' Range("D4").PasteSpecial Transpose:=True
   ' Next l
    
    
End Sub

Private Sub Separate_Click()
  'I am Copying the Data from Sheet1 (Source) to Sheet4 (Destination)
    Sheets("Sheet1").Range("Q5:DZ120000").Copy
    Sheets("Sheet4").Range("K5").PasteSpecial Paste:=xlPasteAllMergingConditionalFormats
    'counts
    Sheets("Sheet1").Range("G2:H3").Copy
    Sheets("Sheet4").Range("E2").PasteSpecial xlPasteValues
    'divisions
    Sheets("Sheet1").Range("K5:O120000").Copy
    Sheets("Sheet4").Range("E5").PasteSpecial xlPasteValues
  
  
  Rows(4).Select
    Selection.Copy
    For i = 17 To 36 Step 13
        Rows(i).Select
        Selection.Insert Shift:=xlDown
        Selection.Copy
    Next i
    
   'One more loop is required i guess
   'For k = 1 To k = 12
    'For j = 18 To 37 Step 13
        'Rows(j).Select
        'Selection.Insert Shift:=xlDown
        'Selection.Copy
   ' Next j
    'Next k
    
        

    'For l = 17 To 36 Step 13

    'Range("D4:D17").Copy
   ' Range("D4").PasteSpecial Transpose:=True
   ' Next l
Sheets("Sheet1").Columns(4).Copy Destination:=Sheets("Sheet4").Columns(2)
    Sheets("Sheet1").Columns(5).Copy Destination:=Sheets("Sheet4").Columns(3)
    
    Dim SelRng As String
    'Dim SelRng1 As String
    Dim Rw_i As Long
    Dim Rw_j As Long
    'Dim i As Long
    Dim j As Long

    Dim x As Long
Dim k As Long
    Dim y As Long



    'For j = 5 To 100 Step 12
    
    For Z = 5 To (Cells(3, 6).Value) + 4 Step 13

'If Cells(Z, 5) = "" Then GoTo lastline


    
    x = Cells(Z, 5).Value
    y = Cells(Z, 6).Value
    
               
   
        SelRng = "B" & x & ":" & "B" & y
        Range(SelRng).Select
        Selection.Copy

        'Rw_i = Range(CStr("E" & j)).Value
        'Rw_j = Range(CStr("F" & j)).Value
       ' SelRng = "B" & Rw_i & ":" & "B" & Rw_j
       ' Range("SelRng").Select
        'Selection.Copy

    'Range("G2:H3").Copy
    'Range("E2").PasteSpecial xlPasteValues

    'SelRng = Range(CStr("K" & i)).Value
    
    'SelRng1 = Range(CStr("M" & j - 2))
    'Range(SelRng1).Select
    'Range("Rw_i").Offset(50, 0).Select
  
    SelRng = "m" & Z - 1
    Range(SelRng).Select
  
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=True, Transpose:=True
     
     
     Next Z
 'Next j
 
 
    For Z = 5 To (Cells(3, 6).Value) + 4 Step 13

'If Cells(Z, 5) = "" Then GoTo lastline


    
    x = Cells(Z, 8).Value
    y = Cells(Z, 9).Value
    
               
   
        SelRng = "B" & x & ":" & "B" & y
        Range(SelRng).Select
        Selection.Copy

        'Rw_i = Range(CStr("E" & j)).Value
        'Rw_j = Range(CStr("F" & j)).Value
       ' SelRng = "B" & Rw_i & ":" & "B" & Rw_j
       ' Range("SelRng").Select
        'Selection.Copy

    'Range("G2:H3").Copy
    'Range("E2").PasteSpecial xlPasteValues

    'SelRng = Range(CStr("K" & i)).Value
    
    'SelRng1 = Range(CStr("M" & j - 2))
    'Range(SelRng1).Select
    'Range("Rw_i").Offset(50, 0).Select
  
    SelRng = "AB" & Z - 1
    Range(SelRng).Select
  
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=True, Transpose:=True
     
     
     Next Z
 'Next j
End Sub

Private Sub Valuesofdistance_Click()
    Sheets("Sheet1").Columns(4).Copy Destination:=Sheets("Sheet4").Columns(2)
    Sheets("Sheet1").Columns(5).Copy Destination:=Sheets("Sheet4").Columns(3)
    
    Dim SelRng As String
    'Dim SelRng1 As String
    Dim Rw_i As Long
    Dim Rw_j As Long
    Dim i As Long
    Dim j As Long

    Dim x As Long
Dim k As Long
    Dim y As Long



    'For j = 5 To 100 Step 12
    
    For Z = 5 To (Cells(3, 6).Value) + 4 Step 13

'If Cells(Z, 5) = "" Then GoTo lastline


    
    x = Cells(Z, 5).Value
    y = Cells(Z, 6).Value
    
               
   
        SelRng = "B" & x & ":" & "B" & y
        Range(SelRng).Select
        Selection.Copy

        'Rw_i = Range(CStr("E" & j)).Value
        'Rw_j = Range(CStr("F" & j)).Value
       ' SelRng = "B" & Rw_i & ":" & "B" & Rw_j
       ' Range("SelRng").Select
        'Selection.Copy

    'Range("G2:H3").Copy
    'Range("E2").PasteSpecial xlPasteValues

    'SelRng = Range(CStr("K" & i)).Value
    
    'SelRng1 = Range(CStr("M" & j - 2))
    'Range(SelRng1).Select
    'Range("Rw_i").Offset(50, 0).Select
  
    SelRng = "m" & Z - 1
    Range(SelRng).Select
  
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=True, Transpose:=True
     
     
     Next Z
 'Next j
 
 
    For Z = 5 To (Cells(3, 6).Value) + 4 Step 13

'If Cells(Z, 5) = "" Then GoTo lastline


    
    x = Cells(Z, 8).Value
    y = Cells(Z, 9).Value
    
               
   
        SelRng = "B" & x & ":" & "B" & y
        Range(SelRng).Select
        Selection.Copy

        'Rw_i = Range(CStr("E" & j)).Value
        'Rw_j = Range(CStr("F" & j)).Value
       ' SelRng = "B" & Rw_i & ":" & "B" & Rw_j
       ' Range("SelRng").Select
        'Selection.Copy

    'Range("G2:H3").Copy
    'Range("E2").PasteSpecial xlPasteValues

    'SelRng = Range(CStr("K" & i)).Value
    
    'SelRng1 = Range(CStr("M" & j - 2))
    'Range(SelRng1).Select
    'Range("Rw_i").Offset(50, 0).Select
  
    SelRng = "AB" & Z - 1
    Range(SelRng).Select
  
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=True, Transpose:=True
     
     
     Next Z
 'Next j

End Sub


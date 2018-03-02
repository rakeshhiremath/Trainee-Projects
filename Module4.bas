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

       
  
    SelRng = "m" & Z - 1
    Range(SelRng).Select
  
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=True, Transpose:=True
     
     
     Next Z

 
 
    For Z = 5 To (Cells(3, 6).Value) + 4 Step 13

'If Cells(Z, 5) = "" Then GoTo lastline


    
    x = Cells(Z, 8).Value
    y = Cells(Z, 9).Value
    
               
   
        SelRng = "B" & x & ":" & "B" & y
        Range(SelRng).Select
        Selection.Copy

     
  
    SelRng = "AB" & Z - 1
    Range(SelRng).Select
  
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=True, Transpose:=True
     
     
     Next Z
 'Next j

End Sub


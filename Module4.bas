Attribute VB_Name = "Module4"
Sub Clear_Click1()
    Range("A4:DZ120000").Clear
    Range("d2:G3").Clear
End Sub
Sub Copycells_Click1()

'I am Copying the Data from Sheet1 (Source) to Sheet4 (Destination)
    Sheets("Sheet1").Range("Q5:DZ120000").Copy
    Sheets("Sheet4").Range("K5").PasteSpecial Paste:=xlPasteAllMergingConditionalFormats
    'counts
    Sheets("Sheet1").Range("F2:H3").Copy
    Sheets("Sheet4").Range("d2").PasteSpecial xlPasteValues
    'divisions
    Sheets("Sheet1").Range("K5:O120000").Copy
    Sheets("Sheet4").Range("E5").PasteSpecial xlPasteValues
    
  End Sub

Sub InsertRows_Click1()


    Rows(4).Select
    Selection.Copy
    
   
    For i = 17 To (Cells(3, 6).Value) Step ((Cells(3, 4).Value) + 2)
        Rows(i).Select
        Selection.Insert Shift:=xlDown
        Selection.Copy
    Next i
    
  
End Sub


Sub Valuesofdistance_Click1()


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

    
    For Z = 6 To (Cells(3, 6).Value) + 4 Step (Cells(3, 4).Value + 2)
 
    x = Cells(Z, 5).Value
    y = Cells(Z, 6).Value
    
        SelRng = "B" & x & ":" & "B" & y
        Range(SelRng).Select
        Selection.Copy
    SelRng = "m" & Z - 1
    Range(SelRng).Select
  
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=True, Transpose:=True
       
     Next Z
 
    For Z = 6 To (Cells(3, 6).Value) + 4 Step (Cells(3, 4).Value + 2)
    
    x = Cells(Z, 8).Value
    y = Cells(Z, 9).Value
    
        SelRng = "B" & x & ":" & "B" & y
        Range(SelRng).Select
        Selection.Copy

    SelRng = "AB" & Z - 1
    Range(SelRng).Select
  
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=True, Transpose:=True
     
     
     Next Z

End Sub

Sub allfunctions_1()
Call Clear_Click1
Call Copycells_Click1
Call InsertRows_Click1
Call Valuesofdistance_Click1
End Sub

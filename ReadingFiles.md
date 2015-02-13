Sub AllFiles()

   Dim Num As Integer
   Num = 1
   
   Dim i As Integer
   Dim iend As Integer
   
   iend = 3                                                                  'Number of files
   
   For i = 1 To iend
   
    Call OpenFile(Num)                                                       'Iterations for each file

    Num = Num + 1
    ActiveCell.Offset(0, Number) = ActiveCell
    Next i

   
End Sub

'---------------------------------------------------------------------------------------------------------------------------

Sub OpenFile(Number As Integer)                                               'Sub for opening a file

    Dim FilePath As String
    
    FilePath = "C:\Users\MMiche01\Desktop\"                                   'File path here (except filename)
    
    Open FilePath & Number & ".501" For Input As #Number                      'I have named the files 1.501, 2.501, 3.501
    
    row_number = 0
    
    Do Until EOF(Number) 
    
        Line Input #Number, LineFromFile                                     'Reads each line
        
        LineItems = Right(LineFromFile, 10)                                  'I want the right 10 characters 
        
        ActiveCell.Offset(row_number, Number - 1).Value = LineItems          'Read next line
        

        
        row_number = row_number + 1
        
    Loop
    

    
    Close #Number
    
End Sub

'-----------------------------------------------------------------------------------------------------------------------

Sub Change()
    Dim lngRow As Long                               'Row Count
    lngRow = Cells(Rows.Count, "A").End(xlUp).Row
    'MsgBox (lngRow)
    Dim x As Integer
    x = lngRow - 1
    

    Dim MyArray() As String                  'Setting var length
    ReDim MyArray(x) As String
    
    Dim StartN As Integer
    Dim EndN As Integer
    EndN = UBound(MyArray)
    
    For StartN = 1 To EndN
        MyArray(StartN) = ActiveCell(StartN, 1).Value
        
    Next StartN
    
    
    Dim OutPut As String
    OutPut = ""
       
    For StartN = 1 To EndN
        OutPut = OutPut & " or frame_id = " & MyArray(StartN)
    Next StartN
    
    Cells(2, 3) = OutPut

End Sub


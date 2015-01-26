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




Sub AllFiles()

   Dim Num As Integer
   Num = 1
   
   Dim i As Integer
   Dim iend As Integer
   
   iend = 3
   
   For i = 1 To iend
   
    Call OpenFile(Num)

    Num = Num + 1
    ActiveCell.Offset(0, Number) = ActiveCell
    Next i

   
End Sub


Sub OpenFile(Number As Integer)

    Dim FilePath As String
    
    FilePath = "C:\Users\MMiche01\Desktop\"
    
    Open FilePath & Number & ".501" For Input As #Number
    
    row_number = 0
    
    Do Until EOF(Number)
    
        Line Input #Number, LineFromFile
        
        LineItems = Right(LineFromFile, 10)
        
        ActiveCell.Offset(row_number, Number - 1).Value = LineItems
        

        
        row_number = row_number + 1
        
    Loop
    

    
    Close #Number
    
End Sub




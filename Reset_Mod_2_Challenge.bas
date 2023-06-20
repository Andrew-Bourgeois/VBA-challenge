Attribute VB_Name = "Module2"
Sub Reset_Exercise()
    ' ------- USE THIS SUBROUTINE TO RESET THE EXERCISE -----
    
    ' alias current worksheet as "crt"
    Dim crt As Worksheet
    
    ' Loop through each worksheet
    For Each crt In Worksheets
    
        crt.Range("I:Q").ClearContents
        crt.Range("I:Q").ClearFormats
    
    Next
    
End Sub

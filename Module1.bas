Attribute VB_Name = "Module1"
Sub probandoLoop()

    Dim i As Integer
    
    i = 1
    
    Do While i <= 10
        If ActiveCell.Value > 10 Then
        ActiveCell.Interior.Color = rgbCoral
        End If
        ActiveCell.Offset(1, 0).Select
        
        i = i + 1
    Loop
    

End Sub

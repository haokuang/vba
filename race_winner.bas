Attribute VB_Name = "Module1"
Sub FindWinner()
    Dim fastest, CurrentTime, CurrentNumber As Integer
    Dim CurrentFirstName, CurrentlastName As String
    
    ActiveSheet.Range("F4").Select
    
    'initialize info
    fastest = ActiveCell.Value
    CurrentFirstName = ActiveCell.Offset(0, -3).Value
    CurrentlastName = ActiveCell.Offset(0, -4).Value
    CurrentNumber = ActiveCell.Offset(0, -5).Value
    
    Do
        ActiveCell.Offset(1, 0).Select
        If IsEmpty(ActiveCell) Then
            Exit Do
        End If
        
        CurrentTime = CInt(ActiveCell.Value)
        If CurrentTime < fastest Then
            fastest = CurrentTime
            CurrentFirstName = ActiveCell.Offset(0, -3).Value
            CurrentlastName = ActiveCell.Offset(0, -4).Value
            CurrentNumber = ActiveCell.Offset(0, -5).Value
        End If
    Loop
    
    MsgBox "The fastest time was " & fastest & " minutes and that was turned in by " & CurrentFirstName & " " & CurrentlastName & " wearing number " & CurrentNumber & "."
    
End Sub

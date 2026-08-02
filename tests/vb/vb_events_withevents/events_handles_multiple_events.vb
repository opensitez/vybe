' vybe-test: vb/vb_events_withevents/events_handles_multiple_events
' origin: languages/vb/tests/vb/test_vb_events_withevents.rs

Class Button
    Public Event Click()
    Public Event MouseEnter()
    
    Public Sub SimulateClick()
        RaiseEvent Click()
    End Sub
    
    Public Sub SimulateHover()
        RaiseEvent MouseEnter()
    End Sub
End Class

Module M
    Private WithEvents Btn As Button
    
    ' One handler method handling multiple events
    Private Sub OnUserInteraction() Handles Btn.Click, Btn.MouseEnter
        Console.WriteLine("Interaction occurred")
    End Sub
    
    Sub Main()
        Btn = New Button()
        Btn.SimulateClick()
        Btn.SimulateHover()
    End Sub
End Module

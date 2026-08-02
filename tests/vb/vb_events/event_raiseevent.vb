' vybe-test: vb/vb_events/event_raiseevent
' origin: languages/vb/tests/vb/test_vb_events.rs

Class Timer
    Public Event Tick()
    Public Sub Fire()
        RaiseEvent Tick()
    End Sub
End Class

Module M
    Sub OnTick()
        Console.WriteLine("tick!")
    End Sub
    Sub Main()
        Dim t As New Timer()
        AddHandler t.Tick, AddressOf OnTick
        t.Fire()
        t.Fire()
    End Sub
End Module

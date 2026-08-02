' vybe-test: vb/vb_custom_event_remove_nonexistent/test_vb_event_remove_duplicate_subscribers_one_by_one
' origin: languages/vb/tests/vb/test_vb_custom_event_remove_nonexistent.rs

Imports System

Class Publisher
    Public Event Tick As Action
    Public Sub Sound()
        RaiseEvent Tick()
    End Sub
End Class

Module Program
    Private Sub OnTick() : Console.WriteLine("Tick") : End Sub

    Sub Main()
        Dim p As New Publisher()
        AddHandler p.Tick, AddressOf OnTick
        AddHandler p.Tick, AddressOf OnTick
        p.Sound()
        Console.WriteLine("---")
        RemoveHandler p.Tick, AddressOf OnTick
        p.Sound()
    End Sub
End Module

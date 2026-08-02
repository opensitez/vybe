' vybe-test: vb/vb_event_handler_add_remove_count/test_vb_event_handler_subscribing_multiple_events_same_handler
' origin: languages/vb/tests/vb/test_vb_event_handler_add_remove_count.rs

Imports System

Class DualSource
    Public Event EventA As EventHandler
    Public Event EventB As EventHandler

    Public Sub TriggerBoth()
        RaiseEvent EventA(Me, EventArgs.Empty)
        RaiseEvent EventB(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim src As New DualSource()
        Dim commonHandler As EventHandler = Sub(s, e) Console.WriteLine("Common Handler Fired")

        AddHandler src.EventA, commonHandler
        AddHandler src.EventB, commonHandler
        src.TriggerBoth()
    End Sub
End Module

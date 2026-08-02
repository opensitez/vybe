' vybe-test: vb/vb_custom_event_remove_nonexistent/test_vb_event_add_handler_during_event_execution
' origin: languages/vb/tests/vb/test_vb_custom_event_remove_nonexistent.rs

Imports System

Class DynamicEmitter
    Public Event Trigger As Action
    Public Sub Fire()
        RaiseEvent Trigger()
    End Sub
End Class

Module Program
    Sub Main()
        Dim de As New DynamicEmitter()
        Dim h2 As Action = Sub() Console.WriteLine("H2 Executed")

        AddHandler de.Trigger, Sub()
            Console.WriteLine("H1 Executing & Adding H2")
            AddHandler de.Trigger, h2
        End Sub

        de.Fire()
        Console.WriteLine("Second Fire:")
        de.Fire()
    End Sub
End Module

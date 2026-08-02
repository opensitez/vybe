' vybe-test: vb/vb_custom_event_thread_sync/test_vb_event_handler_chained_raise_event
' origin: languages/vb/tests/vb/test_vb_custom_event_thread_sync.rs

Imports System

Class EventChain
    Public Event Stage1 As EventHandler
    Public Event Stage2 As EventHandler

    Public Sub StartChain()
        RaiseEvent Stage1(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim ec As New EventChain()
        AddHandler ec.Stage1, Sub(s, e)
            Console.WriteLine("Stage 1")
            ' Raise stage 2 inside stage 1 handler
            AddHandler ec.Stage2, Sub(s2, e2) Console.WriteLine("Stage 2")
        End Sub
        ec.StartChain()
    End Sub
End Module

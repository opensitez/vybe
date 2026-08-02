' vybe-test: vb/vb_custom_event_remove_nonexistent/test_vb_event_remove_nonexistent_handler_no_op
' origin: languages/vb/tests/vb/test_vb_custom_event_remove_nonexistent.rs

Imports System

Class EventSource
    Public Event Action As Action
    Public Sub Fire()
        RaiseEvent Action()
    End Sub
End Class

Module Program
    Private Sub Handler()
        Console.WriteLine("Handler Executed")
    End Sub

    Sub Main()
        Dim src As New EventSource()
        RemoveHandler src.Action, AddressOf Handler
        Console.WriteLine("Remove completed safely")
        src.Fire()
    End Sub
End Module

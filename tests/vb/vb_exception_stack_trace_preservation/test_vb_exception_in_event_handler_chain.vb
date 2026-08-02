' vybe-test: vb/vb_exception_stack_trace_preservation/test_vb_exception_in_event_handler_chain
' origin: languages/vb/tests/vb/test_vb_exception_stack_trace_preservation.rs

Imports System

Class Publisher
    Public Event Fire As Action
    Public Sub RaiseFire()
        RaiseEvent Fire()
    End Sub
End Class

Module Program
    Sub Main()
        Dim p As New Publisher()
        AddHandler p.Fire, Sub() Console.WriteLine("Handler 1")
        AddHandler p.Fire, Sub() Throw New Exception("Handler 2 Failed")
        AddHandler p.Fire, Sub() Console.WriteLine("Handler 3")

        Try
            p.RaiseFire()
        Catch ex As Exception
            Console.WriteLine("Caught in Main: " & ex.Message)
        End Try
    End Sub
End Module

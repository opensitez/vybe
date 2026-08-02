' vybe-test: vb/vb_exception_stack_trace_preservation/test_vb_exception_in_delegate_multicast_stops_chain
' origin: languages/vb/tests/vb/test_vb_exception_stack_trace_preservation.rs

Imports System

Module Program
    Private Sub First()
        Console.WriteLine("First Executed")
        Throw New InvalidOperationException("First Failed")
    End Sub
    Private Sub Second()
        Console.WriteLine("Second Executed")
    End Sub

    Sub Main()
        Dim act As Action = AddressOf First
        act = CType([Delegate].Combine(act, New Action(AddressOf Second)), Action)
        Try
            act()
        Catch ex As Exception
            Console.WriteLine("Caught: " & ex.Message)
        End Try
    End Sub
End Module

' vybe-test: vb/vb_aggregate_exception_handling/test_vb_aggregate_exception_flatten_and_handle
' origin: languages/vb/tests/vb/test_vb_aggregate_exception_handling.rs

Imports System

Module Program
    Sub Main()
        Dim ex1 As New InvalidOperationException("Op1 failed")
        Dim ex2 As New ArgumentException("Arg2 invalid")
        Dim agg As New AggregateException(ex1, ex2)

        Console.WriteLine(agg.InnerExceptions.Count)

        agg.Handle(Function(e)
            Console.WriteLine("Handled: " & e.GetType().Name)
            Return True
        End Function)
    End Sub
End Module

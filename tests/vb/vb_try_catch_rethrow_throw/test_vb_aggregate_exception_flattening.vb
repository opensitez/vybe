' vybe-test: vb/vb_try_catch_rethrow_throw/test_vb_aggregate_exception_flattening
' origin: languages/vb/tests/vb/test_vb_try_catch_rethrow_throw.rs

Imports System
Imports System.Collections.Generic

Module Program
    Sub Main()
        Try
            Dim inner1 As New InvalidOperationException("Op1")
            Dim inner2 As New ArgumentException("Op2")
            Throw New AggregateException("Batch Failed", inner1, inner2)
        Catch ex As AggregateException
            For Each inner In ex.InnerExceptions
                Console.WriteLine("Inner: " & inner.Message)
            Next
        End Try
    End Sub
End Module

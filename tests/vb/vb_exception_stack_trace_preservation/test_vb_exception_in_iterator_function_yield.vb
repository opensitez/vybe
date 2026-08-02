' vybe-test: vb/vb_exception_stack_trace_preservation/test_vb_exception_in_iterator_function_yield
' origin: languages/vb/tests/vb/test_vb_exception_stack_trace_preservation.rs

Imports System
Imports System.Collections.Generic

Module Program
    Private Iterator Function GenerateItems() As IEnumerable(Of Integer)
        Yield 1
        Yield 2
        Throw New InvalidOperationException("Generator Error")
    End Function

    Sub Main()
        Try
            For Each item In GenerateItems()
                Console.WriteLine("Item: " & item)
            Next
        Catch ex As Exception
            Console.WriteLine("Caught in Iterator Loop: " & ex.Message)
        End Try
    End Sub
End Module

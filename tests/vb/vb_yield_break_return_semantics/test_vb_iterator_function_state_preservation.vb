' vybe-test: vb/vb_yield_break_return_semantics/test_vb_iterator_function_state_preservation
' origin: languages/vb/tests/vb/test_vb_yield_break_return_semantics.rs

Imports System.Collections.Generic

Module Program
    Private Iterator Function FibonacciSequence(limit As Integer) As IEnumerable(Of Integer)
        Dim a = 0
        Dim b = 1
        For i As Integer = 1 To limit
            Yield a
            Dim temp = a + b
            a = b
            b = temp
        Next
    End Function

    Sub Main()
        Dim fibs = FibonacciSequence(6)
        Console.WriteLine(String.Join(",", fibs))
    End Sub
End Module

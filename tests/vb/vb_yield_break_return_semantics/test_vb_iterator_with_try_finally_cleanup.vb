' vybe-test: vb/vb_yield_break_return_semantics/test_vb_iterator_with_try_finally_cleanup
' origin: languages/vb/tests/vb/test_vb_yield_break_return_semantics.rs

Imports System.Collections.Generic

Module Program
    Private Iterator Function GeneratorWithFinally() As IEnumerable(Of Integer)
        Try
            Yield 100
            Yield 200
        Finally
            Console.WriteLine("Iterator Finally Executed")
        End Try
    End Function

    Sub Main()
        For Each item In GeneratorWithFinally()
            Console.WriteLine("Item: " & item)
        Next
    End Sub
End Module

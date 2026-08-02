' vybe-test: vb/vb_yield_break_return_semantics/test_vb_iterator_partial_consumption_runs_finally
' origin: languages/vb/tests/vb/test_vb_yield_break_return_semantics.rs

Imports System.Collections.Generic

Module Program
    Private Iterator Function InfiniteWithFinally() As IEnumerable(Of Integer)
        Try
            Dim i = 1
            While True
                Yield i
                i += 1
            End While
        Finally
            Console.WriteLine("Cleaned Up Infinite Generator")
        End Try
    End Function

    Sub Main()
        For Each num In InfiniteWithFinally()
            Console.WriteLine("Got: " & num)
            If num = 2 Then Exit For
        Next
    End Sub
End Module

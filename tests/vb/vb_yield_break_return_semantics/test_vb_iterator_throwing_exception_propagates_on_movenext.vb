' vybe-test: vb/vb_yield_break_return_semantics/test_vb_iterator_throwing_exception_propagates_on_movenext
' origin: languages/vb/tests/vb/test_vb_yield_break_return_semantics.rs

Imports System
Imports System.Collections.Generic

Module Program
    Private Iterator Function FaultyGen() As IEnumerable(Of Integer)
        Yield 1
        Throw New InvalidOperationException("Iterator Fault")
    End Function

    Sub Main()
        Try
            For Each item In FaultyGen()
                Console.WriteLine("Item: " & item)
            Next
        Catch ex As InvalidOperationException
            Console.WriteLine(ex.Message)
        End Try
    End Sub
End Module

' vybe-test: vb/vb_yield_break_return_semantics/test_vb_iterator_recursive_flattening
' origin: languages/vb/tests/vb/test_vb_yield_break_return_semantics.rs

Imports System.Collections.Generic

Module Program
    Private Iterator Function FlattenTree(nodeValue As Integer, depth As Integer) As IEnumerable(Of Integer)
        Yield nodeValue
        If depth > 0 Then
            For Each child In FlattenTree(nodeValue * 10, depth - 1)
                Yield child
            Next
        End If
    End Function

    Sub Main()
        Console.WriteLine(String.Join(",", FlattenTree(1, 2)))
    End Sub
End Module

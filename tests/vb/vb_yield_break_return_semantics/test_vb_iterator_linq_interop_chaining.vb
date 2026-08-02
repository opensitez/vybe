' vybe-test: vb/vb_yield_break_return_semantics/test_vb_iterator_linq_interop_chaining
' origin: languages/vb/tests/vb/test_vb_yield_break_return_semantics.rs

Imports System.Collections.Generic
Imports System.Linq

Module Program
    Private Iterator Function Numbers() As IEnumerable(Of Integer)
        For i As Integer = 1 To 10
            Yield i
        Next
    End Function

    Sub Main()
        Dim evens = Numbers().Where(Function(n) n Mod 2 = 0).Take(3)
        Console.WriteLine(String.Join(",", evens))
    End Sub
End Module

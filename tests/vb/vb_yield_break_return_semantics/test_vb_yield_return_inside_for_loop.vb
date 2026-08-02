' vybe-test: vb/vb_yield_break_return_semantics/test_vb_yield_return_inside_for_loop
' origin: languages/vb/tests/vb/test_vb_yield_break_return_semantics.rs

Imports System.Collections.Generic

Module Program
    Private Iterator Function RangeGenerator(startVal As Integer, count As Integer) As IEnumerable(Of Integer)
        For i As Integer = 0 To count - 1
            Yield startVal + i
        Next
    End Function

    Sub Main()
        Dim items = RangeGenerator(100, 4)
        Console.WriteLine(String.Join("-", items))
    End Sub
End Module

' vybe-test: vb/vb_yield_break_return_semantics/test_vb_yield_return_generic_struct
' origin: languages/vb/tests/vb/test_vb_yield_break_return_semantics.rs

Imports System.Collections.Generic

Structure Pair
    Public Key As String
    Public Val As Integer
End Structure

Module Program
    Private Iterator Function GeneratePairs() As IEnumerable(Of Pair)
        Yield New Pair With {.Key = "K1", .Val = 10}
        Yield New Pair With {.Key = "K2", .Val = 20}
    End Function

    Sub Main()
        For Each p In GeneratePairs()
            Console.WriteLine(p.Key & "=" & p.Val)
        Next
    End Sub
End Module

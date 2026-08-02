' vybe-test: vb/vb_yield_break_return_semantics/test_vb_yield_return_nested_loops
' origin: languages/vb/tests/vb/test_vb_yield_break_return_semantics.rs

Imports System.Collections.Generic

Module Program
    Private Iterator Function GridPoints() As IEnumerable(Of String)
        For r As Integer = 0 To 1
            For c As Integer = 0 To 1
                Yield r & ":" & c
            Next
        Next
    End Function

    Sub Main()
        Console.WriteLine(String.Join(" ", GridPoints()))
    End Sub
End Module

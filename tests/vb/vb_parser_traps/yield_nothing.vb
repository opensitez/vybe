' vybe-test: vb/vb_parser_traps/yield_nothing
' origin: languages/vb/tests/vb/test_vb_parser_traps.rs

Imports System.Collections.Generic

Module M
    Iterator Function Generate() As IEnumerable(Of Object)
        Yield Nothing
    End Function

    Sub Main()
        For Each item In Generate()
            Console.WriteLine(item Is Nothing)
        Next
    End Sub
End Module

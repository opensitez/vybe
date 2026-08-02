' vybe-test: vb/vb_yield_statement/yield_statement_iterator
' origin: languages/vb/tests/vb/test_vb_yield_statement.rs

Imports System.Collections.Generic

Module M
    Iterator Function GetNumbers() As IEnumerable(Of Integer)
        Yield 10
        Yield 20
        Yield 30
    End Function

    Sub Main()
        For Each num In GetNumbers()
            Console.WriteLine(num)
        Next
    End Sub
End Module

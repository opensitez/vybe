' vybe-test: vb/vb_yield_basic/yield_return_basic
' origin: languages/vb/tests/vb/test_vb_yield_basic.rs

Imports System.Collections.Generic

Module M
    Iterator Function GetNumbers() As IEnumerable(Of Integer)
        Yield 1
        Yield 2
        Yield 3
    End Function

    Sub Main()
        For Each n In GetNumbers()
            Console.WriteLine(n)
        Next
    End Sub
End Module

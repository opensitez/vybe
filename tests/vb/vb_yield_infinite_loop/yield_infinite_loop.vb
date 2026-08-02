' vybe-test: vb/vb_yield_infinite_loop/yield_infinite_loop
' origin: languages/vb/tests/vb/test_vb_yield_infinite_loop.rs

Imports System.Collections.Generic
Imports System.Linq

Module M
    Iterator Function Generate() As IEnumerable(Of Integer)
        Dim i = 0
        While True
            Yield i
            i += 1
        End While
    End Function

    Sub Main()
        Dim numbers = Generate().Take(3)
        For Each n In numbers
            Console.WriteLine(n)
        Next
    End Sub
End Module

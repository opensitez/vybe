' vybe-test: vb/vb_iterator_try_finally/iterator_try_finally
' origin: languages/vb/tests/vb/test_vb_iterator_try_finally.rs

Imports System.Collections.Generic

Module M
    Iterator Function Generate() As IEnumerable(Of Integer)
        Try
            Yield 1
            Yield 2
        Finally
            Console.WriteLine("Cleanup")
        End Try
    End Function

    Sub Main()
        For Each num In Generate()
            Console.WriteLine(num)
        Next
    End Sub
End Module

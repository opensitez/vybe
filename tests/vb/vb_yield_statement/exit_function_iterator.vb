' vybe-test: vb/vb_yield_statement/exit_function_iterator
' origin: languages/vb/tests/vb/test_vb_yield_statement.rs

Imports System.Collections.Generic

Module M
    Iterator Function GetNumbersUntil() As IEnumerable(Of Integer)
        Yield 1
        Exit Function ' Terminates iterator
        Yield 2
    End Function

    Sub Main()
        For Each num In GetNumbersUntil()
            Console.WriteLine(num)
        Next
    End Sub
End Module

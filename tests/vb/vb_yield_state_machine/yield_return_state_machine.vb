' vybe-test: vb/vb_yield_state_machine/yield_return_state_machine
' origin: languages/vb/tests/vb/test_vb_yield_state_machine.rs

Imports System.Collections.Generic

Module M
    Iterator Function GetEvenNumbers(max As Integer) As IEnumerable(Of Integer)
        For i As Integer = 1 To max
            If i Mod 2 = 0 Then
                Yield i
            End If
        Next
    End Function

    Sub Main()
        For Each n In GetEvenNumbers(5)
            Console.WriteLine(n)
        Next
    End Sub
End Module

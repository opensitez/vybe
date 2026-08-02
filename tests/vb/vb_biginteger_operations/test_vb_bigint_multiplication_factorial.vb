' vybe-test: vb/vb_biginteger_operations/test_vb_bigint_multiplication_factorial
' origin: languages/vb/tests/vb/test_vb_biginteger_operations.rs

Imports System.Numerics

Module Program
    Sub Main()
        Dim fact As BigInteger = 1
        For i As Integer = 1 To 20
            fact *= i
        Next
        Console.WriteLine(fact.ToString())
    End Sub
End Module

' vybe-test: vb/vb_big_integer_pow_mod_pow/test_vb_big_integer_factorial_simulation
' origin: languages/vb/tests/vb/test_vb_big_integer_pow_mod_pow.rs

Imports System.Numerics

Module Program
    Private Function Factorial(n As Integer) As BigInteger
        Dim result As BigInteger = 1
        For i As Integer = 2 To n
            result *= i
        Next
        Return result
    End Function

    Sub Main()
        Dim fact20 = Factorial(20)
        Console.WriteLine(fact20.ToString())
    End Sub
End Module

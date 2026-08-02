' vybe-test: vb/vb_divide_by_zero_exception_handling/test_vb_integer_mod_by_zero_throws_divide_by_zero_exception
' origin: languages/vb/tests/vb/test_vb_divide_by_zero_exception_handling.rs

Imports System

Module Program
    Sub Main()
        Try
            Dim a As Integer = 10
            Dim b As Integer = 0
            Dim res As Integer = a Mod b
            Console.WriteLine(res)
        Catch ex As DivideByZeroException
            Console.WriteLine("Mod DivideByZeroException Caught")
        End Try
    End Sub
End Module
